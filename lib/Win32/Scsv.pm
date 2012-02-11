package Win32::Scsv;

use strict;
use warnings;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
use Carp;
use Cwd qw(getcwd abs_path);

require Exporter;
our @ISA       = qw(Exporter);
our @EXPORT    = qw();
our @EXPORT_OK = qw(convcsv);
our $VERSION   = '0.01';

sub convcsv {
    # Comment by Klaus Eichner, 11/02/2012
    # I have copied the example code from
    # http://bytes.com/topic/perl/answers/770333-how-convert-csv-file-excel-file
    #
    # and from
    # http://www.tek-tips.com/faqs.cfm?fid=6715

    my ($wbname, $csvname) = @_;

    my $sheetname = '';
    if ($wbname =~ s{% (.*) \z}''xms) {
        $sheetname = $1;
    }

    unless (-f $wbname) {
        croak "Error-0010: file '$wbname' not found";
    }

    unless ($wbname =~ m{\A (.*) \. (xls x?) \z}xmsi) {
        croak "Error-0020: Workbookname: '$wbname' does not have an Excel extension";
    }
    my ($wbstem, $wbext) = ($1, lc($2));

    unless (defined $csvname) {
        if ($sheetname eq '') {
            $csvname = $wbstem.'.csv';
        }
        else {
            $csvname = $sheetname.'.csv';
        }
    }

    $wbname  =~ s{/}'\\'xmsg;
    $csvname =~ s{/}'\\'xmsg;

    # clean out the CSV file (or create an empty CSV file, if it does not exist)
    open my $ofh, '>', $csvname or croak "Error-0030: Can't open > '$csvname' because $!";
    close $ofh;

    my $wbabs  = abs_path($wbname);
    my $csvabs = abs_path($csvname);

    my $curdir = getcwd;
    $curdir =~ s{/}'\\'xmsg;

    my $excel;

    # use existing instance if Excel is already running
    eval { $excel = Win32::OLE->GetActiveObject('Excel.Application') };
    croak "Error-0040: Excel not installed" if $@;

    unless (defined $excel) {
        $excel = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
          or croak "Error-0050: Oops, cannot start Excel";
    }
 
    $excel->{DisplayAlerts}=0;
    my $book = $excel->Workbooks->Open($wbabs);

    my $sheet;
    if ($sheetname eq '') {
        $sheet = $book->Worksheets(1);
        unless (defined $sheet) {
            croak "Error-0060: Can't find first Sheet in Workbook '$wbname'";
        }
    }
    else {
        $sheet = $book->Worksheets($sheetname);
        unless (defined $sheet) {
            croak "Error-0070: Can't find Sheet '$sheetname' in Workbook '$wbname'";
        }
    }

    $sheet->Activate;

    $book->SaveAs($csvabs, xlCSV);
}

1;

__END__

=head1 NAME

Win32::Scsv - Convert Excel file to CSV using Win32::OLE

=head1 SYNOPSIS

    use Win32::Scsv qw(convcsv);

    convcsv('Test Excel File.xlsx%Tabelle3' => 'dummy.csv');
    convcsv('Test Excel File.xlsx%Tabelle Test');

=head1 AUTHOR

Klaus Eichner <klaus03@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2009-2011 by Klaus Eichner

All rights reserved. This program is free software; you can redistribute
it and/or modify it under the terms of the artistic license 2.0,
see http://www.opensource.org/licenses/artistic-license-2.0.php

=cut
