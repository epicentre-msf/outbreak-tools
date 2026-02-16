#!/usr/bin/env perl

# -----------------------------------------------------------------------------
# remove_const_underscores.pl
# Pass a VBA file or directory and the script normalises constant identifiers by
# stripping underscores in their names and synchronising all usages outside
# strings/comments. Designed for predictable refactors when underscores have
# special meaning in downstream tooling.
# -----------------------------------------------------------------------------

use strict;
use warnings;
use File::Find qw(find);
use File::Basename qw(fileparse);

if (!@ARGV) {
    print STDERR "Usage: $0 <file-or-directory> [more paths]\n";
    exit 1;
}

my %stats = (
    files_processed   => 0,
    files_changed     => 0,
    constants_changed => 0,
);

# Allowlist VBA source extensions so we ignore stray binary files in recursive
# runs.
my %allowed_ext = map { $_ => 1 } qw(.bas .cls .frm .ctl .pag .mod .vba .txt);

for my $path (@ARGV) {
    if (-d $path) {
        find(
            {
                wanted => sub {
                    return if -d $_;
                    my $file = $File::Find::name;
                    process_file($file, 0);
                },
                no_chdir => 1
            },
            $path
        );
    } elsif (-f $path) {
        process_file($path, 1);
    } else {
        warn "Skipping $path: not a file or directory\n";
    }
}

if ($stats{files_changed}) {
    print "\nTotal constants renamed: $stats{constants_changed} across $stats{files_changed} file" .
          ($stats{files_changed} == 1 ? "" : "s") . "\n";
} else {
    print "No constant names with underscores were found in the supplied paths.\n";
}

exit 0;

sub process_file {
    my ($file, $report_no_change) = @_;

    my (undef, undef, $ext) = fileparse($file, qr/\.[^\.]+/);
    return unless defined $ext && $allowed_ext{ lc $ext };

    $stats{files_processed}++;

    my $original = read_file($file);
    return unless defined $original;

    my ($updated, $renamed, $usage_updates) = transform_content($original);

    unless ($renamed) {
        print "No changes: $file\n" if $report_no_change;
        return;
    }

    write_file($file, $updated) or return;

    $stats{files_changed}++;
    $stats{constants_changed} += $renamed;

    my @parts;
    push @parts, "$renamed constant" . ($renamed == 1 ? '' : 's') . " renamed";
    push @parts, "$usage_updates usage" . ($usage_updates == 1 ? '' : 's') . " updated" if $usage_updates;

    print "Updated $file (" . join(', ', @parts) . ")\n";
}

sub read_file {
    my ($file) = @_;
    open my $fh, '<:raw', $file or do {
        warn "Failed to read $file: $!\n";
        return;
    };
    local $/;
    my $content = <$fh>;
    close $fh;
    return $content;
}

sub write_file {
    my ($file, $content) = @_;
    open my $fh, '>:raw', $file or do {
        warn "Failed to write $file: $!\n";
        return 0;
    };
    print {$fh} $content;
    close $fh;
    return 1;
}

sub transform_content {
    my ($content) = @_;

    my @chunks = split(/(\r?\n)/, $content, -1);
    my $rebuilt = '';
    my $total_renamed = 0;
    my %renames;

    for (my $i = 0; $i <= $#chunks; $i += 2) {
        my $line = $chunks[$i];
        my $newline = ($i + 1 <= $#chunks) ? $chunks[$i + 1] : '';

        my ($new_line, $renamed) = transform_line($line, \%renames);
        $rebuilt .= $new_line . $newline;
        $total_renamed += $renamed;
    }

    # After declarations are cleaned, sweep the module again to rewrite usages
    # with the renamed identifiers.
    my ($final_content, $usage_updates) = apply_identifier_replacements($rebuilt, \%renames);

    return ($final_content, $total_renamed, $usage_updates);
}

sub transform_line {
    my ($line, $rename_map) = @_;
    return ($line, 0) unless defined $line && $line =~ /const/i;

    my ($code, $comment) = split_code_and_comment($line);
    my ($new_code, $renamed) = rename_constants_in_code($code, $rename_map);

    return ($new_code . $comment, $renamed);
}

sub split_code_and_comment {
    my ($line) = @_;
    my $len = length $line;
    my $in_string = 0;
    my $lower = lc $line;

    for (my $i = 0; $i < $len; $i++) {
        my $char = substr($line, $i, 1);

        if ($char eq '"') {
            my $next = substr($line, $i + 1, 1);
            if (defined $next && $next eq '"') {
                $i++;
                next;
            }
            $in_string = !$in_string;
            next;
        }

        next if $in_string;

        if ($char eq '\'') {
            return (substr($line, 0, $i), substr($line, $i));
        }

        if ($i <= $len - 3 && substr($lower, $i, 3) eq 'rem') {
            my $prev = $i > 0 ? substr($line, $i - 1, 1) : '';
            my $next = $i + 3 < $len ? substr($line, $i + 3, 1) : '';
            if ($prev !~ /[A-Za-z0-9_]/ && ($next eq '' || $next =~ /\s|[:']/)) {
                return (substr($line, 0, $i), substr($line, $i));
            }
        }
    }

    return ($line, '');
}

sub rename_constants_in_code {
    my ($code, $rename_map) = @_;
    return ($code, 0) if $code eq '';

    my $result = '';
    my $cursor = 0;
    my $total = 0;

    while ($code =~ /\bConst\b/ig) {
        my $start = $-[0];
        my $end = $+[0];
        $result .= substr($code, $cursor, $end - $cursor);

        my $tail = substr($code, $end);
        my ($processed, $consumed, $renamed) = process_const_tail($tail, $rename_map);
        $result .= $processed;

        $cursor = $end + $consumed;
        $total += $renamed;
        pos($code) = $cursor;
    }

    $result .= substr($code, $cursor);
    return ($result, $total);
}

sub process_const_tail {
    my ($text, $rename_map) = @_;
    my $len = length $text;
    my $i = 0;
    my $state = 'expect_name';
    my $paren_depth = 0;
    my $in_string = 0;
    my $output = '';
    my $renamed = 0;

    while ($i < $len) {
        my $char = substr($text, $i, 1);

        if (!$in_string && $char eq ':') {
            last;
        }

        if ($state eq 'expect_name') {
            if ($char =~ /\s/) {
                $output .= $char;
                $i++;
                next;
            }

            if ($char eq '_') {
                $output .= $char;
                $i++;
                next;
            }

            if ($char eq '[') {
                my $j = $i + 1;
                while ($j < $len && substr($text, $j, 1) ne ']') {
                    $j++;
                }
                if ($j >= $len) {
                    $output .= $char;
                    $i++;
                    next;
                }
                my $name_body = substr($text, $i + 1, $j - $i - 1);
                my $new_body = $name_body;
                if ($new_body =~ /_/) {
                    my $compressed = $new_body;
                    $compressed =~ s/_+//g;
                    $new_body = $compressed eq '' ? $name_body : $compressed;
                }
                if ($new_body ne $name_body) {
                    # Track bracketed identifiers so later replacements stay in
                    # sync with declarations that use attribute syntax.
                    $renamed++;
                    $rename_map->{'[' . $name_body . ']'} = '[' . $new_body . ']';
                }
                $output .= '[' . $new_body . ']';
                $i = $j + 1;
                $state = 'body';
                next;
            }

            if ($char =~ /[A-Za-z_]/) {
                my $j = $i + 1;
                while ($j < $len && substr($text, $j, 1) =~ /[A-Za-z0-9_]/) {
                    $j++;
                }
                my $suffix = '';
                if ($j < $len && substr($text, $j, 1) =~ /[%&@!#\$]/) {
                    $suffix = substr($text, $j, 1);
                    $j++;
                }
                my $base = substr($text, $i, $j - $i - length $suffix);
                my $new_base = $base;
                if ($new_base =~ /_/) {
                    my $compressed = $new_base;
                    $compressed =~ s/_+//g;
                    $new_base = $compressed eq '' ? $base : $compressed;
                }
                if ($new_base ne $base) {
                    # Record the renamed identifier (including any type suffix)
                    # so we can update usages once we finish parsing this line.
                    $renamed++;
                    my $original = $base . $suffix;
                    my $updated = $new_base . $suffix;
                    $rename_map->{$original} = $updated;
                }
                $output .= $new_base . $suffix;
                $i = $j;
                $state = 'body';
                next;
            }

            $output .= $char;
            $i++;
            next;
        }

        $output .= $char;

        if ($char eq '"') {
            if ($in_string) {
                my $next = substr($text, $i + 1, 1);
                if (defined $next && $next eq '"') {
                    $output .= $next;
                    $i += 2;
                    next;
                }
                $in_string = 0;
                $i++;
                next;
            }
            $in_string = 1;
            $i++;
            next;
        }

        if (!$in_string) {
            if ($char eq '(') {
                $paren_depth++;
            } elsif ($char eq ')' && $paren_depth > 0) {
                $paren_depth--;
            } elsif ($char eq ',' && $paren_depth == 0) {
                $state = 'expect_name';
            }
        }

        $i++;
    }

    return ($output, $i, $renamed);
}

sub apply_identifier_replacements {
    my ($content, $renames) = @_;
    return ($content, 0) unless %{$renames};

    my @keys = sort { length($b) <=> length($a) } keys %{$renames};
    # Longest-first matching prevents shorter identifiers from hijacking a
    # longer, overlapping name (e.g. CONST and CONSTONE).
    my @chunks = split(/(\r?\n)/, $content, -1);
    my $rebuilt = '';
    my $total_updates = 0;

    for (my $i = 0; $i <= $#chunks; $i += 2) {
        my $line = $chunks[$i];
        my $newline = ($i + 1 <= $#chunks) ? $chunks[$i + 1] : '';

        my ($code, $comment) = split_code_and_comment($line);
        my ($new_code, $count) = replace_identifiers($code, $renames, \@keys);

        $rebuilt .= $new_code . $comment . $newline;
        $total_updates += $count;
    }

    return ($rebuilt, $total_updates);
}

sub replace_identifiers {
    my ($code, $renames, $ordered_keys) = @_;
    return ($code, 0) if $code eq '';

    my $result = '';
    my $len = length $code;
    my $i = 0;
    my $in_string = 0;
    my $replaced = 0;

    while ($i < $len) {
        my $char = substr($code, $i, 1);

        if ($char eq '"') {
            $result .= $char;
            my $next = substr($code, $i + 1, 1);
            if ($in_string && defined $next && $next eq '"') {
                $result .= $next;
                $i += 2;
                next;
            }
            $in_string = !$in_string;
            $i++;
            next;
        }

        if (!$in_string) {
            # Attempt to replace identifiers only when we're in code territory.
            my $matched = 0;
            for my $old (@{$ordered_keys}) {
                my $length = length $old;
                next if $i + $length > $len;
                next unless substr($code, $i, $length) eq $old;

                my $prev = $i > 0 ? substr($code, $i - 1, 1) : '';
                my $next_char = $i + $length < $len ? substr($code, $i + $length, 1) : '';

                next if $prev =~ /[A-Za-z0-9_]/;

                my $last_char = substr($old, -1);
                if ($last_char =~ /[A-Za-z0-9_]/ && $next_char =~ /[A-Za-z0-9_]/) {
                    next;
                }

                $result .= $renames->{$old};
                $i += $length;
                $matched = 1;
                $replaced++;
                last;
            }
            next if $matched;
        }

        $result .= $char;
        $i++;
    }

    return ($result, $replaced);
}
