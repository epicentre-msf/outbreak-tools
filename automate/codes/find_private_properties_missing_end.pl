#!/usr/bin/env perl

use strict;
use warnings;

use Cwd qw(abs_path);
use File::Basename qw(dirname);
use File::Find ();
use File::Path qw(make_path);
use File::Spec;
use Getopt::Long qw(GetOptions);

my $root   = '.';
my $output = 'temp/private_properties_missing_end.txt';

GetOptions(
    'root=s'   => \$root,
    'output=s' => \$output,
) or die "Usage: $0 [--root DIR] [--output FILE]\n";

$root = abs_path($root // '.')
  or die "Unable to resolve root path\n";
die "Root directory '$root' does not exist\n" unless -d $root;

my $output_dir = dirname($output);
if (defined $output_dir && length $output_dir && $output_dir ne '.') {
    make_path($output_dir) unless -d $output_dir;
}

open my $out_fh, '>', $output
  or die "Failed to open output file '$output': $!\n";

my @candidates;
File::Find::find(
    {
        wanted => sub {
            return unless -f $_;
            return unless $_ =~ /\.(?:cls|bas|frm)$/i;
            push @candidates, $File::Find::name;
        },
        no_chdir => 1,
    },
    $root
);

@candidates = sort @candidates;

my $found_missing = 0;
for my $file (@candidates) {
    open my $fh, '<', $file or do {
        warn "Unable to read '$file': $!\n";
        next;
    };

    my $relative_path = File::Spec->abs2rel($file, $root);
    my @open_properties;
    my $line_number = 0;

    while (my $line = <$fh>) {
        $line_number++;

        my $code = $line;
        $code =~ s/\s+'[^\n\r]*$//;     # strip trailing inline comment
        next if $code =~ /^\s*(?:'|Rem\b)/i;

        if ($code =~ /\bPrivate\s+Property\b/i) {
            push @open_properties,
              {
                line => $line_number,
                text => $line,
              };
        }

        if ($code =~ /\bEnd\s+Property\b/i && @open_properties) {
            pop @open_properties;
        }
    }

    close $fh;

    if (@open_properties) {
        $found_missing = 1;
        for my $property (@open_properties) {
            chomp(my $line_text = $property->{text});
            print {$out_fh} "$relative_path:$property->{line}: $line_text\n";
        }
    }
}

print {$out_fh} "No missing 'End Property' statements found.\n"
  unless $found_missing;

close $out_fh;

exit $found_missing ? 1 : 0;
