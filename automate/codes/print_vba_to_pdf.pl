#!/usr/bin/env perl

# -----------------------------------------------------------------------------
# print_vba_to_pdf.pl
# Export VBA modules as a syntax-highlighted PDF. The script gathers files,
# normalises their line endings into temporary copies, runs pygmentize for HTML
# + CSS, then calls wkhtmltopdf to create a nicely formatted document. Designed
# for VS Code tasks so we can share readable snippets without touching source.
# -----------------------------------------------------------------------------

use strict;
use warnings;
use Getopt::Long qw(GetOptions);
use File::Find qw(find);
use File::Spec;
use File::Path qw(make_path);
use File::Temp qw(tempdir);
use IPC::Open3;
use Symbol qw(gensym);
use Cwd qw(abs_path);

my %MODULE_EXT = map { $_ => 1 } qw(.bas .cls .frm .ctl .pag .mod);  # VBA source extensions
my $DEFAULT_STYLE = 'friendly';

sub usage {
    die "Usage: $0 --output <file.pdf> [--title <title>] [--style <pygments-style>] [--recursive] [--pygmentize <path>] [--wkhtmltopdf <path>] <paths...>\n";
}

my $output;
my $recursive   = 0;
my $title       = 'VBA Modules';
my $style       = $DEFAULT_STYLE;
my $pygmentize  = $ENV{PYGMENTIZE}  || 'pygmentize';   # allow callers to override tool paths
my $wkhtmltopdf = $ENV{WKHTMLTOPDF} || 'wkhtmltopdf';  # otherwise rely on PATH discovery

GetOptions(
    'output=s'      => \$output,
    'recursive'     => \$recursive,
    'title=s'       => \$title,
    'style=s'       => \$style,
    'pygmentize=s'  => \$pygmentize,
    'wkhtmltopdf=s' => \$wkhtmltopdf,
) or usage();

usage() unless defined $output;
my @targets = @ARGV;
usage() unless @targets;

my @modules = gather_modules(\@targets, $recursive);
if (!@modules) {
    warn "No VBA modules found for the supplied path(s).\n";
    exit 1;
}

ensure_tool([$pygmentize, '-V'], 'pygmentize');
ensure_tool([$wkhtmltopdf, '--version'], 'wkhtmltopdf');

my $tmpdir = tempdir(CLEANUP => 1);
my $css = capture([$pygmentize, '-S', $style, '-f', 'html']);
my @sections;
for my $module (@modules) {
    # Each module gets copied to the temp dir with Unix newlines so pygmentize
    # doesn't inherit Windows-style CRLF artefacts.
    my $normalized = normalize_module($module, $tmpdir);
    my $snippet = capture([
        $pygmentize,
        '-f', 'html',
        '-O', 'linenos=table',
        '-l', 'vbnet',
        $normalized,
    ]);
    push @sections, sprintf("<section>\n<h2>%s</h2>\n%s\n</section>", html_escape($module), $snippet);
}

my $body   = join("\n\n", @sections);
my $escaped_title = html_escape($title);
# Inline HTML+CSS keeps wkhtmltopdf self-contained (no external assets).
my $html = <<"HTML";
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>$escaped_title</title>
<style>
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; padding: 2rem; font-size: 14pt; }
code, pre { font-size: 12pt; }
section { margin-bottom: 2rem; }
h1, h2 { font-weight: 600; }
$css
</style>
</head>
<body>
<h1>$escaped_title</h1>
$body
</body>
</html>
HTML

my $output_abs = File::Spec->rel2abs($output);
my ($out_vol, $out_dir, $out_file) = File::Spec->splitpath($output_abs);
my $dest_dir = File::Spec->catpath($out_vol, $out_dir, '');
$dest_dir =~ s{[\/]+$}{};
if (defined $dest_dir && length $dest_dir) {
    make_path($dest_dir) unless -d $dest_dir;
}

my $html_path = File::Spec->catfile($tmpdir, 'vba_modules.html');
write_file($html_path, $html);
run_command([
    $wkhtmltopdf,
    '--quiet',
    '--margin-left',  '15mm',
    '--margin-right', '15mm',
    '--margin-top',   '20mm',
    '--margin-bottom','20mm',
    $html_path,
    $output_abs,   # final PDF output
]);

print "Created $output_abs\n";
exit 0;

sub gather_modules {
    my ($targets, $recursive) = @_;
    my %seen;
    for my $target (@$targets) {
        if (-f $target) {
            # Path is a single file; queue it if the extension matches.
            add_if_module(\%seen, $target);
        } elsif (-d $target) {
            if ($recursive) {
                # Dive into subdirectories when the recursive flag is set.
                find(
                    {
                        wanted => sub {
                            return if -d $_;
                            add_if_module(\%seen, $File::Find::name);
                        },
                        no_chdir => 1,
                    },
                    $target
                );
            } else {
                opendir(my $dh, $target) or do {
                    warn "warning: failed to read directory $target: $!\n";
                    next;
                };
                while (my $entry = readdir $dh) {
                    next if $entry eq '.' || $entry eq '..';
                    my $full = File::Spec->catfile($target, $entry);
                    next unless -f $full;
                    # Non-recursive mode inspects only the top-level contents.
                    add_if_module(\%seen, $full);
                }
                closedir $dh;
            }
        } else {
            warn "warning: $target not found\n";
        }
    }
    return sort keys %seen;
}

sub add_if_module {
    my ($set, $candidate) = @_;
    my $ext = lc(extname($candidate));
    return unless $MODULE_EXT{$ext};
    my $abs = File::Spec->rel2abs($candidate);
    # Store canonical absolute path so duplicates across inputs collapse.
    $set->{$abs} = 1;
}

sub extname {
    my ($path) = @_;
    return '' unless defined $path;
    $path =~ /(\.[^.\\\/]+)$/;
    return $1 // '';
}

sub ensure_tool {
    my ($cmd, $label) = @_;
    # Probe external dependencies (pygmentize/wkhtmltopdf) so the failure is
    # caught early with a readable message.
    my $stderr = gensym;
    my $pid = open3(undef, my $out, $stderr, @$cmd);
    while (<$out>) { }
    close $out;
    my $err = do { local $/; <$stderr> // '' };
    close $stderr;
    waitpid($pid, 0);
    my $code = $? >> 8;
    if ($code != 0) {
        die "$label command failed or not found. $err";
    }
}

sub capture {
    my ($cmd) = @_;
    # Run a command, returning stdout and including stderr in any exception.
    my $stderr = gensym;
    my $pid = open3(undef, my $out, $stderr, @$cmd);
    local $/;
    my $stdout = <$out> // '';
    close $out;
    my $err = <$stderr> // '';
    close $stderr;
    waitpid($pid, 0);
    my $code = $? >> 8;
    if ($code != 0) {
        die "Command '@{[join ' ', @$cmd]}' failed: $err";
    }
    return $stdout;
}

my $NORMALIZED_SEQ = 0;
sub normalize_module {
    my ($path, $tmpdir) = @_;
    # Preserve the original file by writing a numbered copy in the working dir.
    open my $in, '<:raw', $path or die "Failed to read $path: $!\n";
    local $/;
    my $content = <$in> // '';
    close $in;
    # Replace CRLF/CR with LF so rendered HTML respects line breaks.
    $content =~ s/\r\n/\n/g;
    $content =~ s/\r/\n/g;
    my $filename = sprintf('module_%04d.tmp', ++$NORMALIZED_SEQ);
    my $normalized_path = File::Spec->catfile($tmpdir, $filename);
    open my $out, '>:raw', $normalized_path or die "Failed to write $normalized_path: $!\n";
    print {$out} $content;
    close $out;
    return $normalized_path;
}

sub html_escape {
    my ($text) = @_;
    return '' unless defined $text;
    $text =~ s/&/&amp;/g;
    $text =~ s/</&lt;/g;
    $text =~ s/>/&gt;/g;
    $text =~ s/"/&quot;/g;
    return $text;
}

sub write_file {
    my ($path, $content) = @_;
    open my $fh, '>:utf8', $path or die "Failed to write $path: $!\n";
    print {$fh} $content;
    close $fh;
}

sub run_command {
    my ($cmd) = @_;
    system(@$cmd) == 0
        or die "Command '@{[join ' ', @$cmd]}' failed with exit code " . ($? >> 8) . "\n";
}
