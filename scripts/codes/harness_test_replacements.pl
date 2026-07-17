#!/usr/bin/env perl

use strict;
use warnings;
use List::Util qw(max);

################################################################################
# harness_test_replacements.pl
# ------------------------------------------------------------------------------
# This script mirrors the logic implemented in automate/codes/harness_test_replacements.R
# but is written in Perl for maintainability.  It ingests a VBA module, scrubs out the
# legacy Rubberduck harness boilerplate, and injects the CustomTest harness conventions
# used by this project.  The script deliberately favors readability over terseness; the
# heavy commenting is meant to capture the intent of every transformation so future
# maintainers can update the routine with confidence.
################################################################################

#------------------------------------------------------------------------------
# Utility helpers
#------------------------------------------------------------------------------

# insert_after(
#     \@lines,    # array reference of the current module content (1 entry per line)
#     $index,     # zero-based index after which new lines should be inserted
#     \@payload   # array reference with the lines we wish to inject
# )
#
# The helper gracefully handles the two edge cases encountered when translating
# the R helper append_after(): inserting at the start of the file and appending
# to the end.  No-op when there is nothing to insert.
sub insert_after {
    my ($lines_ref, $index, $payload_ref) = @_;
    return if !$payload_ref || !@{$payload_ref};

    if ($index < 0) {
        # Prepend when append_after in R received an index <= 0.
        @{$lines_ref} = (@{$payload_ref}, @{$lines_ref});
        return;
    }

    if ($index >= $#{$lines_ref}) {
        # Append when the index is at or beyond the current upper bound.
        push @{$lines_ref}, @{$payload_ref};
        return;
    }

    # General case: splice the payload immediately after the requested index.
    splice @{$lines_ref}, $index + 1, 0, @{$payload_ref};
}

# find_sub_end(
#     \@lines,  # module lines
#     $start    # index of the Sub signature line
# ) -> $index of the matching End Sub (or the final line when missing)
#
# The routine walks forward from the signature until it encounters "End Sub".
# While the source modules are expected to be well formed, the defensive
# fallback to the last line keeps the script resilient if a module is partially
# edited.
sub find_sub_end {
    my ($lines_ref, $start) = @_;
    return $#{$lines_ref} if !defined $start || $start < 0;

    for my $idx ($start + 1 .. $#{$lines_ref}) {
        return $idx if $lines_ref->[$idx] =~ /^\s*End\s+Sub\b/i;
    }

    return $#{$lines_ref};
}

# escape_vba_string($text) -> string
# Escapes embedded double quotes so the generated VBA string literal remains valid.
sub escape_vba_string {
    my ($text) = @_;
    return '' if !defined $text;
    $text =~ s/"/""/g;
    return $text;
}

#------------------------------------------------------------------------------
# Entry point argument handling
#------------------------------------------------------------------------------

my $target_path = shift @ARGV
    or die "Pass the VBA test module path (relative or absolute) as the first argument.\n";

-f $target_path
    or die sprintf("File not found: %s\n", $target_path);

# Read the module while normalising line endings to \n for easier manipulation.
open my $fh, '<', $target_path or die sprintf("Unable to open %s: %s\n", $target_path, $!);
binmode $fh;
local $/;
my $raw = <$fh> // '';
close $fh;

$raw =~ s/\r\n/\n/g;   # convert CRLF -> LF
$raw =~ s/\r/\n/g;      # and stray CR -> LF
my @lines = split /\n/, $raw;

#------------------------------------------------------------------------------
# Discover the module name for later use when stamping the harness
#------------------------------------------------------------------------------

my $module_name = '';
for my $idx (0 .. $#lines) {
    if ($lines[$idx] =~ /^Attribute\s+VB_Name\s*=\s*"([^"]+)"/) {
        $module_name = $1;
        last;
    }
}

#------------------------------------------------------------------------------
# Normalise module annotations for CustomTest harness
#------------------------------------------------------------------------------

my $folder_insert_idx;
my @filtered_lines;
for my $idx (0 .. $#lines) {
    if ($lines[$idx] =~ /^\s*'\s*\@TestModule\b/i) {
        $folder_insert_idx //= $idx;
        next;
    }
    push @filtered_lines, $lines[$idx];
}
@lines = @filtered_lines;

my $has_custom_folder = scalar grep { /^(?:\s*'\s*\@Folder\s*\(\s*"CustomTests"\s*\))/i } @lines;
if (!$has_custom_folder) {
    my $insert_idx;
    if (defined $folder_insert_idx) {
        $insert_idx = $folder_insert_idx;
    } else {
        ($insert_idx) = grep { $lines[$_] =~ /^Option\s+Explicit\b/i } 0 .. $#lines;
        if (defined $insert_idx) {
            $insert_idx += 1;
        } else {
            ($insert_idx) = grep { $lines[$_] =~ /^Attribute\s+VB_Name\b/i } 0 .. $#lines;
            $insert_idx = defined $insert_idx ? $insert_idx + 1 : 0;
        }
    }

    my $folder_line = '\'@Folder("CustomTests")';
    $insert_idx = 0 if !defined $insert_idx;
    $insert_idx = scalar @lines if $insert_idx > @lines;
    splice @lines, $insert_idx, 0, $folder_line;
}

#------------------------------------------------------------------------------
# Remove legacy harness bootstrap and normalise harness references
#------------------------------------------------------------------------------

@lines = grep { $_ !~ /CreateObject\("Rubberduck\.AssertClass"\)/ } @lines;
@lines = grep { $_ !~ /CreateObject\("Rubberduck\.FakesProvider"\)/ } @lines;
@lines = grep { $_ !~ /^Option\s+Private\s+Module/i } @lines;

for my $line (@lines) {
    $line =~ s/\bPrivate\s+Assert\s+As\s+Object\b/Private Assert As ICustomTest/g;
    $line =~ s/\bDim\s+Assert\s+As\s+Object\b/Dim Assert As ICustomTest/g;
    $line =~ s/\bPrivate\s+Harness\s+As\s+ICustomTest\b/Private Assert As ICustomTest/g;
    $line =~ s/\bDim\s+Harness\s+As\s+ICustomTest\b/Dim Assert As ICustomTest/g;
    $line =~ s/Set\s+Harness\s*=\s*/Set Assert = /g;
    $line =~ s/FailUnexpectedError\s+Harness/FailUnexpectedError Assert/g;
    $line =~ s/Harness\./Assert./g;
    $line =~ s/\bHarness\b/Assert/g;
}

@lines = grep { $_ !~ /Assert\.SetTestName/ } @lines;
@lines = grep { $_ !~ /Assert\.SetTestSubtitle/ } @lines;

@lines = grep { $_ !~ /Private\s+Const\s+OUTPUT_SHEET_NAME/i } @lines;
for my $line (@lines) {
    $line =~ s/OUTPUT_SHEET_NAME/TEST_OUTPUT_SHEET/g;
}

# Ensure the TEST_OUTPUT_SHEET constant exists (insert after Option Explicit when absent).
if (!grep { /TEST_OUTPUT_SHEET/ } @lines) {
    my $anchor = max(-1, (map { $lines[$_] =~ /Option\s+Explicit/i ? $_ : () } 0 .. $#lines));
    my @const_block = ('', 'Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"', '');
    insert_after(\@lines, $anchor, \@const_block);
}

#------------------------------------------------------------------------------
# Harmonise assertion helpers and failure logging semantics
#------------------------------------------------------------------------------

for my $line (@lines) {
    $line =~ s/Assert\.IsNotNothing\s*\(([^,]+),\s*([^)]*)\)/"Assert.ObjectExists($1, \"Object\", $2)"/eg;
    $line =~ s/Assert\.(Fail|Inconclusive)\s*\(/Assert.LogFailure(/g;
    $line =~ s/Assert\.(Fail|Inconclusive)\s+/Assert.LogFailure /g;
    $line =~ s/FailUnexpectedError\s+Assert\s*,\s*"([^"]+)"/"CustomTestLogFailure Assert, \"" . escape_vba_string($1) . "\", Err.Number, Err.Description"/eg;
    $line =~ s/Assert\.PrintResults\s+OUTPUT_SHEET_NAME/Assert.PrintResults TEST_OUTPUT_SHEET/g;
    $line =~ s/Assert\.FlushCurrentTest\s*\([^)]*\)/Assert.Flush/g;
    $line =~ s/Assert\.FlushCurrentTest\b/Assert.Flush/g;
}

#------------------------------------------------------------------------------
# Normalise ModuleInitialize to spin up the CustomTest harness
#------------------------------------------------------------------------------

my ($module_init_idx) = grep { $lines[$_] =~ /^\s*(?:Private|Public)\s+Sub\s+ModuleInitialize\b/i } 0 .. $#lines;
if (defined $module_init_idx) {
    my $module_init_end = find_sub_end(\@lines, $module_init_idx);
    my @body_range = $module_init_idx + 1 .. $module_init_end - 1;

    my @create_calls = grep { $lines[$_] =~ /CustomTest\.Create/ } @body_range;
    if (!@create_calls) {
        my @busy_lines = grep { $lines[$_] =~ /BusyApp/ } @body_range;
        my $insert_after = @busy_lines ? max(@busy_lines) : $module_init_idx;

        my @init_lines = ('    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)');
        if (defined $module_name && length $module_name) {
            push @init_lines, sprintf('    Assert.SetModuleName "%s"', escape_vba_string($module_name));
        }
        insert_after(\@lines, $insert_after, \@init_lines);
    }
    else {
        for my $idx (@create_calls) {
            $lines[$idx] =~ s/CustomTest\.Create\([^)]*\)/CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)/g;
        }
    }
}

#------------------------------------------------------------------------------
# Guarantee ModuleCleanup prints the accumulated results
#------------------------------------------------------------------------------

my ($module_cleanup_idx) = grep { $lines[$_] =~ /^\s*(?:Private|Public)\s+Sub\s+ModuleCleanup\b/i } 0 .. $#lines;
if (defined $module_cleanup_idx) {
    my $module_cleanup_end = find_sub_end(\@lines, $module_cleanup_idx);
    my @body_range = $module_cleanup_idx + 1 .. $module_cleanup_end - 1;

    my $has_print = scalar grep { $lines[$_] =~ /Assert\.PrintResults/ } @body_range;
    if (!$has_print) {
        my @on_error = grep { $lines[$_] =~ /On\s+Error/i } @body_range;
        my $insert_after = @on_error ? max(@on_error) : $module_cleanup_idx;
        my @block = (
            '    If Not Assert Is Nothing Then',
            '        Assert.PrintResults TEST_OUTPUT_SHEET',
            '    End If'
        );
        insert_after(\@lines, $insert_after, \@block);
    }
}

#------------------------------------------------------------------------------
# Ensure TestCleanup flushes the active checking
#------------------------------------------------------------------------------

my ($test_cleanup_idx) = grep { $lines[$_] =~ /^\s*(?:Private|Public)\s+Sub\s+TestCleanup\b/i } 0 .. $#lines;
if (defined $test_cleanup_idx) {
    my $test_cleanup_end = find_sub_end(\@lines, $test_cleanup_idx);
    my @body_range = $test_cleanup_idx + 1 .. $test_cleanup_end - 1;

    my $has_flush = scalar grep { $lines[$_] =~ /Assert\.Flush(?:CurrentTest)?\b/ } @body_range;
    if (!$has_flush) {
        my $insert_after = $test_cleanup_idx;
        my @block = (
            '    If Not Assert Is Nothing Then',
            '        Assert.Flush',
            '    End If'
        );
        insert_after(\@lines, $insert_after, \@block);
    }
} else {
    my ($test_init_idx) = grep { $lines[$_] =~ /^\s*(?:Private|Public)\s+Sub\s+TestInitialize\b/i } 0 .. $#lines;
    my $anchor_idx = defined $test_init_idx ? find_sub_end(\@lines, $test_init_idx)
                     : (defined $module_cleanup_idx ? find_sub_end(\@lines, $module_cleanup_idx)
                        : $#lines);
    my @block = (
        '',
        'Private Sub TestCleanup()',
        '    If Not Assert Is Nothing Then',
        '        Assert.Flush',
        '    End If',
        'End Sub'
    );
    insert_after(\@lines, $anchor_idx, \@block);
    ($test_cleanup_idx) = grep { $lines[$_] =~ /^\s*(?:Private|Public)\s+Sub\s+TestCleanup\b/i } 0 .. $#lines;
}

#------------------------------------------------------------------------------
# Process each @TestMethod annotation to inject CustomTestSetTitles metadata
#------------------------------------------------------------------------------

my @test_annotations = grep { $lines[$_] =~ /^'\@TestMethod/ } 0 .. $#lines;
if (@test_annotations) {
    for my $annotation_idx (reverse @test_annotations) {
        my $annotation = $lines[$annotation_idx];
        my $title = '';
        if (index($annotation, '"') != -1) {
            ($title) = $annotation =~ /\(\"([^\"]*)\"/;
        }
        $title = 'Tests' if !defined($title) || $title eq '';

        my $sig_idx;
        for my $candidate ($annotation_idx + 1 .. $#lines) {
            if ($lines[$candidate] =~ /^\s*(?:Private|Public)\s+Sub\s+[A-Za-z0-9_]+/) {
                $sig_idx = $candidate;
                last;
            }
        }
        next if !defined $sig_idx;

        my $sig_line = $lines[$sig_idx];
        my ($leading_ws) = $sig_line =~ /^(\s*)/;
        my ($method_name) = $sig_line =~ /^\s*(?:Private|Public)\s+Sub\s+([A-Za-z0-9_]+)/;
        $method_name //= '';

        # Promote Private Subs to Public to match Rubberduck expectations.
        $lines[$sig_idx] =~ s/^(\s*)Private(\s+Sub)/$1Public$2/;
        # Guarantee a consistent indent (default to four spaces when none captured).
        my $indent = defined($leading_ws) && $leading_ws ne '' ? $leading_ws : '    ';
        my $escaped_method = escape_vba_string($method_name);

        my $sub_end = find_sub_end(\@lines, $sig_idx);
        my $next_line = ($sig_idx + 1 <= $#lines) ? $lines[$sig_idx + 1] : '';
        if ($next_line !~ /CustomTestSetTitles/) {
            my $escaped_title = escape_vba_string($title);
            my $title_line = sprintf('%sCustomTestSetTitles Assert, "%s", "%s"', $indent, $escaped_title, $escaped_method);
            insert_after(\@lines, $sig_idx, [$title_line]);
        }

        for my $line_idx ($sig_idx .. $sub_end) {
            $lines[$line_idx] =~ s/(CustomTestLogFailure\s+Assert,\s*")([^"]*)(".*)/$1$escaped_method$3/;
        }
    }
}

#------------------------------------------------------------------------------
# Persist the transformed module.  The VBA editor expects CRLF line endings, so
# we restore them and ensure the file terminates with a newline.
#------------------------------------------------------------------------------

my $output = join("\r\n", @lines) . "\r\n";
open my $out, '>', $target_path or die sprintf("Unable to write %s: %s\n", $target_path, $!);
binmode $out;
print {$out} $output;
close $out;

printf "Updated %s\n", $target_path;

exit 0;
