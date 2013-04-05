package Win32::Scsv;

use strict;
use warnings;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
use Carp;
use Cwd qw(getcwd abs_path);
use File::Copy;

require Exporter;
our @ISA       = qw(Exporter);
our @EXPORT    = qw();
our @EXPORT_OK = qw(convcsv convxls);
our $VERSION   = '0.02';

# Comment by Klaus Eichner, 11/02/2012
#
# I have copied the example code from
# http://bytes.com/topic/perl/answers/770333-how-convert-csv-file-excel-file
#
# and from
# http://www.tek-tips.com/faqs.cfm?fid=6715
#
# ...also an excellent source of information with regards to Win32::Ole / Excel is the
# perlmonks-article ("Using Win32::OLE and Excel - Tips and Tricks") at the following site:
# http://www.perlmonks.org/bare/?node_id=153486
#
# In that perlmonks-article there is a link to another article
# ("The Perl Journal #10, courtesy of Jon Orwant")
# http://search.cpan.org/~gsar/libwin32-0.191/OLE/lib/Win32/OLE/TPJ.pod

sub convcsv {
    my ($xls_name, $xls_snumber) = $_[0] =~ m{\A ([^%]*) % ([^%]*) \z}xms ? ($1, $2) : ($_[0], 1);
    my $csv_name = $_[1];

    unless ($xls_name =~ m{\A (.*) \. (xls x?) \z}xmsi) {
        croak "Error-0010: xls_name '$xls_name' does not have an Excel extension (*.xls, *.xlsx)";
    }

    my ($xls_stem, $xls_ext) = ($1, lc($2));

    unless (-f $xls_name) {
        croak "Error-0020: xls_name '$xls_name' not found";
    }

    # remove the CSV file (if it exists)
    if (-e $csv_name) {
        unlink $csv_name or croak "Error-0030: Can't unlink csv_name '$csv_name' because $!";
    }

    # no create a new, empty CSV file (...so that abs_path($csv_name) does not fail...)
    open my $ofh, '>', $csv_name or croak "Error-0040: Can't open > '$csv_name' because $!";
    close $ofh;

    my $xls_abs = abs_path($xls_name); $xls_abs =~ s{/}'\\'xmsg;
    my $csv_abs = abs_path($csv_name); $csv_abs =~ s{/}'\\'xmsg;

    my $ole_excel = get_excel()
      or croak "Error-0050: Can't start Excel";

    my $xls_book = $ole_excel->Workbooks->Open($xls_abs)
       or croak "Error-0060: Can't Workbooks->Open xls_abs '$xls_abs'";

    my $xls_sheet = $xls_book->Worksheets($xls_snumber)
       or croak "Error-0070: Can't find Sheet '$xls_snumber' in xls_abs '$xls_abs'";

    $xls_sheet->Activate;

    $xls_book->SaveAs($csv_abs, xlCSV);
}

sub convxls {
    my ($xls_name, $xls_snumber) = $_[1] =~ m{\A ([^%]*) % ([^%]*) \z}xms ? ($1, $2) : ($_[1], 1);
    my $csv_name = $_[0];
    my $tpl_name = $_[2] && defined($_[2]{'tpl'}) ? $_[2]{'tpl'}    : '';
    my @col_size = $_[2] && defined($_[2]{'csz'}) ? @{$_[2]{'csz'}} : ();

    if ($tpl_name eq '') {
        croak "Error-0210: tpl_name is empty";
    }

    my ($xls_stem, $xls_ext) = $xls_name =~ m{\A (.*) \. (xls x?) \z}xmsi ? ($1, lc($2)) :
      croak "Error-0220: xls_name '$xls_name' does not have an Excel extension of the right type (*.xls, *.xlsx)";

    my ($tpl_stem, $tpl_ext) = $tpl_name =~ m{\A (.*) \. (xls x?) \z}xmsi ? ($1, lc($2)) :
      croak "Error-0230: tpl_name '$tpl_name' does not have an Excel extension of the right type (*.xls, *.xlsx)";

    unless ($xls_ext eq $tpl_ext) {
        croak "Error-0240: extensions do not match between ".
          "xls and tpl ('$xls_ext', '$tpl_ext'), name is ('$xls_name', '$tpl_name')";
    }

    # remove the XLS file (if it exists)
    if (-e $xls_name) {
        unlink $xls_name or croak "Error-0250: Can't unlink xls_name '$xls_name' because $!";
    }

    copy $tpl_name, $xls_name
      or croak "Error-0260: Can't copy tpl_name to xls_name ('$tpl_name', '$xls_name') because $!";

    my $xls_abs = abs_path($xls_name); $xls_abs =~ s{/}'\\'xmsg;
    my $tpl_abs = abs_path($tpl_name); $tpl_abs =~ s{/}'\\'xmsg;
    my $csv_abs = abs_path($csv_name); $csv_abs =~ s{/}'\\'xmsg;

    unless (-f $csv_abs) {
        croak "Error-0270: csv_abs '$csv_abs' not found";
    }

    unless (-f $tpl_abs) {
        croak "Error-0280: tpl_abs '$tpl_abs' not found";
    }

    my $ole_excel = get_excel()
      or croak "Error-0290: Can't start Excel";

    my $xls_book = $ole_excel->Workbooks->Open($xls_abs)
       or croak "Error-0300: Can't Workbooks->Open xls_abs '$xls_abs'";

    my $xls_sheet = $xls_book->Worksheets($xls_snumber)
       or croak "Error-0310: Can't find Sheet '$xls_snumber' in xls_abs '$xls_abs'";

    my $csv_book = $ole_excel->Workbooks->Open($csv_abs)
       or croak "Error-0320: Can't Workbooks->Open csv_abs '$csv_abs'";

    my $csv_sheet = $csv_book->Worksheets(1)
       or croak "Error-0330: Can't find Sheet #1 in csv_abs '$csv_abs'";

    $xls_sheet->Activate; # "...->Activate" is necessary in order to allow "...Range('A1')->Select" later to be effective
    $xls_sheet->Cells->ClearContents;

    $csv_sheet->Cells->Copy;

    $xls_sheet->Range('A1')->PasteSpecial(xlPasteValues);
    $xls_sheet->Cells->EntireColumn->AutoFit;
    $xls_sheet->Columns($_->[0].':'.$_->[0])->{ColumnWidth} = $_->[1] for @col_size;
    $xls_sheet->Range('A1')->Select;

    $xls_book->Save;
}

sub get_excel {
    my $appl;

    # use existing instance if Excel is already running
    eval { $appl = Win32::OLE->GetActiveObject('Excel.Application') };
    return if $@;

    unless (defined $appl) {
        $appl = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
          or return;
    }
 
    $appl->{DisplayAlerts} = 0;

    return $appl;
}

1;

__END__

=head1 NAME

Win32::Scsv - Convert Excel file to CSV using Win32::OLE

=head1 SYNOPSIS

    use Win32::Scsv qw(convcsv convxls);

    convcsv('Test Excel File.xlsx%Tabelle3' => 'dummy.csv');
    convcsv('Test Excel File.xlsx%Tabelle Test');

    convxls('dummy.csv' => 'New.xls%Tab9',
      {'tpl' => 'Template.xls', 'csz' => [['H' => 13.71], ['O' => 3]]});

=head1 AUTHOR

Klaus Eichner <klaus03@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2009-2011 by Klaus Eichner

All rights reserved. This program is free software; you can redistribute
it and/or modify it under the terms of the artistic license 2.0,
see http://www.opensource.org/licenses/artistic-license-2.0.php

=cut
