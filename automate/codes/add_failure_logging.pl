#!/usr/bin/env perl

use strict;
use warnings;

################################################################################
# add_failure_logging.pl
# ------------------------------------------------------------------------------
# Injects the CustomTestLogFailure error-handling pattern into VBA test methods
# that lack it.  For every @TestMethod-annotated Sub that does not already
# contain a TestFail error handler, the script:
#
#   1. Inserts  On Error GoTo TestFail  after the Sub signature (or after the
#      existing CustomTestSetTitles call when present).
#   2. Inserts  Exit Sub  +  TestFail:  +  CustomTestLogFailure  immediately
#      before the final  End Sub.
#
# Usage:
#   perl automate/codes/add_failure_logging.pl <path-to-test.bas>
#
# The file is modified in place.  Line endings are normalised to CRLF so the
# output is VBA-editor-ready.
#
# Designed to be safe and idempotent:
#   - Already-instrumented methods are left untouched.
#   - Multi-line Sub signatures (VBA underscore continuation) are handled.
#   - Methods with an existing On Error GoTo <label> are skipped entirely to
#     avoid conflicting error handlers.
################################################################################

#------------------------------------------------------------------------------
# Utility helpers (same API as harness_test_replacements.pl for consistency)
#------------------------------------------------------------------------------

# insert_after(\@lines, $index, \@payload)
# Splice @payload into @lines immediately after $index.
sub insert_after {
    my ($lines_ref, $index, $payload_ref) = @_;
    return if !$payload_ref || !@{$payload_ref};

    if ($index < 0) {
        @{$lines_ref} = (@{$payload_ref}, @{$lines_ref});
        return;
    }

    if ($index >= $#{$lines_ref}) {
        push @{$lines_ref}, @{$payload_ref};
        return;
    }

    splice @{$lines_ref}, $index + 1, 0, @{$payload_ref};
}

# find_sub_end(\@lines, $start) -> index of the matching End Sub
sub find_sub_end {
    my ($lines_ref, $start) = @_;
    return $#{$lines_ref} if !defined $start || $start < 0;

    for my $idx ($start + 1 .. $#{$lines_ref}) {
        return $idx if $lines_ref->[$idx] =~ /^\s*End\s+Sub\b/i;
    }

    return $#{$lines_ref};
}

# find_signature_end(\@lines, $sig_idx) -> index of the last physical line of
# a possibly multi-line Sub signature (VBA uses trailing underscore for
# continuation).
sub find_signature_end {
    my ($lines_ref, $sig_idx) = @_;
    my $current = $sig_idx;

    while ($current < $#{$lines_ref} && $lines_ref->[$current] =~ /_\s*$/) {
        $current++;
    }

    return $current;
}

# escape_vba_string($text) -> string with embedded double quotes escaped.
sub escape_vba_string {
    my ($text) = @_;
    return '' if !defined $text;
    $text =~ s/"/""/g;
    return $text;
}

#------------------------------------------------------------------------------
# Entry point
#------------------------------------------------------------------------------

my $target_path = shift @ARGV
    or die "Usage: perl add_failure_logging.pl <path-to-test.bas>\n";

-f $target_path
    or die sprintf("File not found: %s\n", $target_path);

# Read and normalise line endings.
open my $fh, '<', $target_path or die sprintf("Unable to open %s: %s\n", $target_path, $!);
binmode $fh;
local $/;
my $raw = <$fh> // '';
close $fh;

$raw =~ s/\r\n/\n/g;
$raw =~ s/\r/\n/g;
my @lines = split /\n/, $raw;

#------------------------------------------------------------------------------
# Discover all @TestMethod annotations
#------------------------------------------------------------------------------

my @test_annotations = grep { $lines[$_] =~ /^'\@TestMethod/ } 0 .. $#lines;

if (!@test_annotations) {
    printf "No \@TestMethod annotations found in %s -- nothing to do.\n", $target_path;
    exit 0;
}

#------------------------------------------------------------------------------
# Track how many methods were instrumented for the summary line.
#------------------------------------------------------------------------------

my $instrumented_count = 0;
my $skipped_count      = 0;

#------------------------------------------------------------------------------
# Process each @TestMethod in REVERSE order so that splice offsets from earlier
# annotations remain valid when we insert lines into later ones first.
#------------------------------------------------------------------------------

for my $annotation_idx (reverse @test_annotations) {
    # Locate the Sub signature that follows the annotation.
    my $sig_idx;
    for my $candidate ($annotation_idx + 1 .. $#lines) {
        if ($lines[$candidate] =~ /^\s*(?:Private|Public)\s+Sub\s+([A-Za-z0-9_]+)/i) {
            $sig_idx = $candidate;
            last;
        }
        # Stop searching if we hit another annotation or a blank gap larger than
        # two lines (defensive: annotation must be immediately above the Sub).
        last if $lines[$candidate] =~ /^'\@/;
    }
    next if !defined $sig_idx;

    # Extract the method name from the signature.
    my ($method_name) = $lines[$sig_idx] =~ /^\s*(?:Private|Public)\s+Sub\s+([A-Za-z0-9_]+)/i;
    $method_name //= 'UnknownTest';

    # Find where the Sub body starts (after any continuation lines).
    my $sig_end  = find_signature_end(\@lines, $sig_idx);
    my $sub_end  = find_sub_end(\@lines, $sig_idx);
    my @body_idx = ($sig_end + 1 .. $sub_end - 1);

    #--------------------------------------------------------------------------
    # Skip when the method already contains failure-logging infrastructure.
    #--------------------------------------------------------------------------

    # Check 1: Already has a TestFail label or CustomTestLogFailure call.
    my $has_test_fail = grep {
        $lines[$_] =~ /^\s*TestFail\s*:/i ||
        $lines[$_] =~ /CustomTestLogFailure/i
    } @body_idx;

    if ($has_test_fail) {
        $skipped_count++;
        next;
    }

    # Check 2: Already has an On Error GoTo pointing to any label (user may
    # have a custom error handler -- do not interfere).
    my $has_on_error_goto = grep {
        $lines[$_] =~ /^\s*On\s+Error\s+GoTo\s+\w/i
    } @body_idx;

    if ($has_on_error_goto) {
        $skipped_count++;
        next;
    }

    #--------------------------------------------------------------------------
    # Determine where to insert "On Error GoTo TestFail".
    #
    # Preferred position: immediately after an existing CustomTestSetTitles call
    # (so the test title is registered before the error handler kicks in).
    # Fallback: immediately after the Sub signature.
    #--------------------------------------------------------------------------

    my ($titles_idx) = grep { $lines[$_] =~ /CustomTestSetTitles/i } @body_idx;
    my $on_error_insert_after = defined $titles_idx ? $titles_idx : $sig_end;

    #--------------------------------------------------------------------------
    # Determine indentation.  Prefer copying the indent of the first non-blank
    # body line; fall back to four spaces.
    #--------------------------------------------------------------------------

    my $indent = '    ';
    for my $idx (@body_idx) {
        if ($lines[$idx] =~ /^(\s+)\S/) {
            $indent = $1;
            last;
        }
    }

    #--------------------------------------------------------------------------
    # Step 1: Insert the error-handling tail BEFORE End Sub.
    #
    #     <blank line>
    #     Exit Sub
    # TestFail:
    #     CustomTestLogFailure Assert, "<method>", Err.Number, Err.Description
    # End Sub          <-- already exists
    #--------------------------------------------------------------------------

    # Recompute sub_end because nothing has been inserted yet.
    $sub_end = find_sub_end(\@lines, $sig_idx);

    my $escaped_method = escape_vba_string($method_name);

    # Check whether Exit Sub already exists right before End Sub (skip if so to
    # avoid duplicating it -- user might have manually added one).
    my $prev_line_idx = $sub_end - 1;
    while ($prev_line_idx > $sig_end && $lines[$prev_line_idx] =~ /^\s*$/) {
        $prev_line_idx--;
    }
    my $has_exit_sub = ($prev_line_idx > $sig_end && $lines[$prev_line_idx] =~ /^\s*Exit\s+Sub\b/i);

    my @tail;
    if (!$has_exit_sub) {
        @tail = (
            '',
            "${indent}Exit Sub",
            "TestFail:",
            "${indent}CustomTestLogFailure Assert, \"${escaped_method}\", Err.Number, Err.Description",
        );
    } else {
        @tail = (
            "TestFail:",
            "${indent}CustomTestLogFailure Assert, \"${escaped_method}\", Err.Number, Err.Description",
        );
    }

    # Insert the tail just before End Sub.
    splice @lines, $sub_end, 0, @tail;

    #--------------------------------------------------------------------------
    # Step 2: Insert "On Error GoTo TestFail" after the chosen anchor.
    #--------------------------------------------------------------------------

    insert_after(\@lines, $on_error_insert_after, ["${indent}On Error GoTo TestFail"]);

    $instrumented_count++;
}

#------------------------------------------------------------------------------
# Persist the result with CRLF line endings.
#------------------------------------------------------------------------------

my $output = join("\r\n", @lines) . "\r\n";
open my $out, '>', $target_path or die sprintf("Unable to write %s: %s\n", $target_path, $!);
binmode $out;
print {$out} $output;
close $out;

printf "Updated %s: %d method(s) instrumented, %d skipped (already handled).\n",
       $target_path, $instrumented_count, $skipped_count;

exit 0;
