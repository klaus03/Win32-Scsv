package Win32::Scsv;

use strict;
use warnings;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Carp;
use File::Spec;
use File::Copy;

require Exporter;
our @ISA       = qw(Exporter);
our @EXPORT    = qw();
our @EXPORT_OK = qw(
  xls_2_csv csv_2_xls xls_2_vbs slurp_vbs import_vbs_book empty_xls
  get_xver get_book get_last_row get_last_col tmp_book open_excel
);

our $VERSION = '0.13';

my $OpenXMLWorkbook = 51; # xlOpenXMLWorkbook

my $vtfalse = Variant(VT_BOOL, 0);
my $vttrue  = Variant(VT_BOOL, 1);

my $ole_global;

my $excel_exe;

for my $office ('', '11', '12', '14') {
    for my $x86 ('', ' (x86)') {
        my $Rn = 'C:\Program Files'.$x86.
          '\Microsoft Office\OFFICE'.$office.'\EXCEL.EXE';

        $excel_exe = $Rn if -f $Rn;
    }
}

sub open_excel {
    unless (defined $excel_exe) {
        croak "Can't find EXCEL.EXE";
    }

    system qq{start /min cmd.exe /k ""$excel_exe" "$_[0]" || pause & exit"};
}

# Comment by Klaus Eichner, 11-Feb-2012:
# **************************************
#
# I have copied the sample code from
# http://bytes.com/topic/perl/answers/770333-how-convert-csv-file-excel-file
#
# ...and from
# http://www.tek-tips.com/faqs.cfm?fid=6715
#
# ...also an excellent source of information with regards to Win32::Ole / Excel is the
# perlmonks-article ("Using Win32::OLE and Excel - Tips and Tricks") at the following site:
# http://www.perlmonks.org/bare/?node_id=153486
#
# ...In that perlmonks-article there is a link to another article
# ("The Perl Journal #10, courtesy of Jon Orwant")
# http://search.cpan.org/~gsar/libwin32-0.191/OLE/lib/Win32/OLE/TPJ.pod
#
# ...I found the following site to identify the different Excel versions (12.0 -> 2007, 11.0 -> 2003, etc...):
# http://www.mrexcel.com/forum/excel-questions/357733-visual-basic-applications-test-finding-excel-version.html
#
# ...I found the following blog ('robhammond.co') to extract Excel macros -- see below subroutine xls_2_vbs()...
# http://robhammond.co/blog/export-vba-code-from-excel-files-using-perl/
#
# ...in this blog ('robhammond.co'), the following 3 additional links were mentioned:
# http://www.perlmonks.org/?node_id=927532
# http://www.perlmonks.org/?node_id=953718
# http://access.mvps.org/access/general/gen0022.htm

# Comment by Klaus Eichner, 12-Jan-2014:
# **************************************
#
# I have copied sample code for import_vbs_file() from
# http://www.mrexcel.com/articles/copy-vba-module.php

sub get_xver {
    my $ole_excel = get_excel() or croak "Can't start Excel";

    my $ver = $ole_excel->Version;
    my $prd =
      $ver eq '14.0' ? '2010' :
      $ver eq '12.0' ? '2007' :
      $ver eq '11.0' ? '2003' :
      $ver eq '10.0' ? '2002' :
      $ver eq  '9.0' ? '2000' :
      $ver eq  '8.0' ? '1997' :
      $ver eq  '7.0' ? '1995' : '????';

    return ($ver, $prd) if wantarray;
    return $ver;
}

sub xls_2_csv {
    my ($xls_name, $xls_snumber) = $_[0] =~ m{\A ([^%]*) % ([^%]*) \z}xms ? ($1, $2) : ($_[0], 1);
    my $csv_name = $_[1];

    unless ($xls_name =~ m{\A (.*) \. (xls x?) \z}xmsi) {
        croak "xls_name '$xls_name' does not have an Excel extension (*.xls, *.xlsx)";
    }

    my ($xls_stem, $xls_ext) = ($1, lc($2));

    unless (-f $xls_name) {
        croak "xls_name '$xls_name' not found";
    }

    my $xls_abs = File::Spec->rel2abs($xls_name); $xls_abs =~ s{/}'\\'xmsg;
    my $csv_abs = File::Spec->rel2abs($csv_name); $csv_abs =~ s{/}'\\'xmsg;

    # remove the CSV file (if it exists)
    if (-e $csv_abs) {
        unlink $csv_abs or croak "Can't unlink csv_abs '$csv_abs' because $!";
    }

    my $ole_excel = get_excel() or croak "Can't start Excel";

    my $xls_book = $ole_excel->Workbooks->Open($xls_abs)
       or croak "Can't Workbooks->Open xls_abs '$xls_abs'";

    my $xls_sheet = $xls_book->Worksheets($xls_snumber)
       or croak "Can't find Sheet '$xls_snumber' in xls_abs '$xls_abs'";

    $xls_sheet->{'Visible'} = $vttrue;

    $xls_sheet->Activate;
    $xls_book->SaveAs($csv_abs, xlCSV);
    $xls_book->Close;
}

sub csv_2_xls {
    my ($xls_name, $xls_snumber) = $_[1] =~ m{\A ([^%]*) % ([^%]*) \z}xms ? ($1, $2) : ($_[1], 1);
    my $csv_name = $_[0];

    my $tpl_name   = $_[2] && defined($_[2]{'tpl'})  ? $_[2]{'tpl'}    : '';
    my @col_size   = $_[2] && defined($_[2]{'csz'})  ? @{$_[2]{'csz'}} : ();
    my @col_fmt    = $_[2] && defined($_[2]{'fmt'})  ? @{$_[2]{'fmt'}} : ();
    my $sheet_prot = $_[2] && defined($_[2]{'prot'}) ? $_[2]{'prot'}   : 0;

    my $init_new = 0;

    if ($tpl_name eq '*') {
        $init_new = 1;
        $tpl_name = '';
    }

    my ($xls_stem, $xls_ext) = $xls_name =~ m{\A (.*) \. (xls x?) \z}xmsi ? ($1, lc($2)) :
      croak "xls_name '$xls_name' does not have an Excel extension of the right type (*.xls, *.xlsx)";

    my $xls_format = $xls_ext eq 'xls' ? xlNormal : $OpenXMLWorkbook;

    my ($tpl_stem, $tpl_ext) =
      $tpl_name eq ''                            ? ('', '')     :
      $tpl_name =~ m{\A (.*) \. (xls x?) \z}xmsi ? ($1, lc($2)) :
      croak "tpl_name '$tpl_name' does not have an Excel extension of the right type (*.xls, *.xlsx)";

    unless ($tpl_name eq '' or $tpl_ext eq $xls_ext) {
        croak "extensions do not match between ".
          "xls and tpl ('$xls_ext', '$tpl_ext'), name is ('$xls_name', '$tpl_name')";
    }

    my $xls_abs = $xls_name eq '' ? '' : File::Spec->rel2abs($xls_name); $xls_abs =~ s{/}'\\'xmsg;
    my $tpl_abs = $tpl_name eq '' ? '' : File::Spec->rel2abs($tpl_name); $tpl_abs =~ s{/}'\\'xmsg;
    my $csv_abs = $csv_name eq '' ? '' : File::Spec->rel2abs($csv_name); $csv_abs =~ s{/}'\\'xmsg;

    if ($init_new) {
        if (-e $xls_abs) {
            unlink $xls_abs or croak "Can't unlink '$xls_abs' because $!";
        }

        my $tmp_ole = get_excel() or croak "Can't start Excel (tmp)";
        my $tmp_book = $tmp_ole->Workbooks->Add or croak "Can't Workbooks->Add xls_abs '$xls_abs' (tmp)";
        $tmp_book->SaveAs($xls_abs, $xls_format);
        $tmp_book->Close;
    }

    if ($tpl_name eq '') {
        unless (-f $xls_name) {
            croak "xls_name ('$xls_name') does not exist and template was not specified";
        }
    }
    else {
        unlink $xls_name;
        copy $tpl_name, $xls_name
          or croak "Can't copy tpl_name to xls_name ('$tpl_name', '$xls_name') because $!";
    }

    unless ($csv_abs eq '' or -f $csv_abs) {
        croak "csv_abs '$csv_abs' not found";
    }

    unless ($tpl_abs eq '' or -f $tpl_abs) {
        croak "tpl_abs '$tpl_abs' not found";
    }

    my $ole_excel = get_excel() or croak "Can't start Excel (new)";
    my $xls_book  = $ole_excel->Workbooks->Open($xls_abs) or croak "Can't Workbooks->Open xls_abs '$xls_abs'";
    my $xls_sheet = $xls_book->Worksheets($xls_snumber) or croak "Can't find Sheet '$xls_snumber' in xls_abs '$xls_abs'";

    $xls_sheet->Activate; # "...->Activate" is necessary in order to allow "...Range('A1')->Select" later to be effective
    $xls_sheet->Unprotect; # unprotect the sheet in any case...
    $xls_sheet->Columns($_->[0])->{NumberFormat} = $_->[1] for @col_fmt;

    unless ($csv_abs eq '') {
        my $csv_book  = $ole_excel->Workbooks->Open($csv_abs) or croak "Can't Workbooks->Open csv_abs '$csv_abs'";
        my $csv_sheet = $csv_book->Worksheets(1) or croak "Can't find Sheet #1 in csv_abs '$csv_abs'";

        $xls_sheet->Cells->ClearContents;
        $csv_sheet->Cells->Copy;
        $xls_sheet->Range('A1')->PasteSpecial(xlPasteValues);
        $xls_sheet->Cells->EntireColumn->AutoFit;

        $csv_book->Close;
    }

    $xls_sheet->Columns($_->[0])->{ColumnWidth}  = $_->[1] for @col_size;

    #~ http://www.mrexcel.com/forum/excel-questions/275645-identifying-freeze-panes-position-sheet-using-visual-basic-applications.html 
    #~ The command "$ole_excel->ActiveWindow->Panes($pi)->VisibleRange->Address"  has currently no use,
    #~ but you never know what it might be good for in the future...
    #~
    #~ Deb-0010: PCount = 4
    #~ Deb-0020: Pane 1 = '$A$1:$E$1'
    #~ Deb-0020: Pane 2 = '$F$1:$AA$1'
    #~ Deb-0020: Pane 3 = '$A$45:$E$102'
    #~ Deb-0020: Pane 4 = '$F$45:$AA$102'
    #~
    #~ print "Deb-0010: PCount = ", $ole_excel->ActiveWindow->Panes->Count, "\n";
    #~ for my $pi (1..$ole_excel->ActiveWindow->Panes->Count) {
    #~     print "Deb-0020: Pane $pi = '", $ole_excel->ActiveWindow->Panes($pi)->VisibleRange->Address, "'\n";
    #~ }
    #~
    #~ However, "FreezePanes", "ScrollRow", "ScrollColumn" and "VisibleRange" are more useful...
    #~
    #~ print "Deb-0030: FreezePanes      = '", $ole_excel->ActiveWindow->FreezePanes,          "'\n";
    #~ print "Deb-0040: ScrollRow        = '", $ole_excel->ActiveWindow->ScrollRow,            "'\n";
    #~ print "Deb-0050: ScrollColumn     = '", $ole_excel->ActiveWindow->ScrollColumn,         "'\n";
    #~ print "Deb-0060: VisibleRange     = '", $ole_excel->ActiveWindow->VisibleRange,         "'\n";
    #~ print "Deb-0070: VisibleRange-Row = '", $ole_excel->ActiveWindow->VisibleRange->Row,    "'\n";
    #~ print "Deb-0070: VisibleRange-Col = '", $ole_excel->ActiveWindow->VisibleRange->Column, "'\n";
    #~
    #~ $ole_excel->ActiveWindow->VisibleRange->Select;

    #~ http://stackoverflow.com/questions/3232920/how-can-i-programmatically-freeze-the-top-row-of-an-excel-worksheet-in-excel-200
    #~ Dim r As Range
    #~ Set r = ActiveCell
    #~ Range("A2").Select
    #~ With ActiveWindow
    #~     .FreezePanes = False
    #~     .ScrollRow = 1
    #~     .ScrollColumn = 1
    #~     .FreezePanes = True
    #~     .ScrollRow = r.Row
    #~ End With
    #~ r.Select

    # Be aware: Even if we try to set ActiveWindow->{ScrollColumn}/{ScrollRow} to "1", this might not succeed,
    # because of frozen panes in the active window. As a consequence, ActiveWindow->{ScrollColumn}/{ScrollRow}
    # could in fact be a value that differs from the original value "1". (this is reflected in the two variables
    # $pos_row/$pos_col).

    $ole_excel->ActiveWindow->{ScrollColumn} = 1;
    $ole_excel->ActiveWindow->{ScrollRow}    = 1;

    my $pos_row = $ole_excel->ActiveWindow->{ScrollRow};
    my $pos_col = $ole_excel->ActiveWindow->{ScrollColumn};

    $xls_sheet->Cells($pos_row, $pos_col)->Select;

    if ($sheet_prot) {
        $xls_sheet->Protect({
          DrawingObjects => $vttrue, 
          Contents       => $vttrue, 
          Scenarios      => $vttrue,
        });
    }

    $xls_book->SaveAs($xls_abs, $xls_format); # ...always use SaveAs(), never use Save() here ...
    $xls_book->Close;
}

sub xls_2_vbs {
    my ($xls_name, $vbs_name) = @_;

    my $list = slurp_vbs($xls_name);

    open my $ofh, '>', $vbs_name or croak "Can't write to '$vbs_name' because $!";

    for my $l (@$list) {
        print {$ofh} "' **>> ", '=' x 50, "\n";
        print {$ofh} "' **>> ", 'Module: ', $l->{'NAME'}, "\n";
        print {$ofh} "' **>> ", '=' x 50, "\n";
        print {$ofh} $l->{'CODE'}, "\n";
        print {$ofh} "' **>> ", '-' x 50, "\n";
    }

    close $ofh;
}

sub slurp_vbs {
    my ($xls_name) = @_;

    my $xls_book  = get_book($xls_name);

    my $xls_proj   = $xls_book->{VBProject}    or croak "Can't create object 'VBProject'";
    my $xls_vbcomp = $xls_proj->{VBComponents} or croak "Can't create object 'VBComponents'";

    my $mlist = [];

    for my $xls_cele (in $xls_vbcomp) {
        my $modname = $xls_cele->Name // '?';
        my $xls_vb  = $xls_cele->{CodeModule}
          or croak "Can't create object 'CodeModule' for modname '$modname'";

        my $lcount = $xls_vb->{CountOfLines};

        if ($lcount) {
            my $body = join '', $xls_vb->Lines(1, $lcount);
            $body =~ s{\r}''xmsg; # fix superfluous linefeeds 
            push @$mlist, { 'NAME' => $modname, 'CODE' => $body };
        }
    }

    $xls_book->Close;

    return $mlist;
}

sub import_vbs_book {
    my ($xls_book, $vbs_name) = @_;

    my $vbs_abs = File::Spec->rel2abs($vbs_name); $vbs_abs =~ s{/}'\\'xmsg;

    my $xls_proj   = $xls_book->{VBProject}    or croak "Can't create object 'VBProject'";
    my $xls_vbcomp = $xls_proj->{VBComponents} or croak "Can't create object 'VBComponents'";

    $xls_vbcomp->Import($vbs_abs);
}

sub empty_xls {
    my $xls_name = $_[0];

    my ($xls_stem, $xls_ext) = $xls_name =~ m{\A (.*) \. (xls x?) \z}xmsi ? ($1, lc($2)) :
      croak "xls_name '$xls_name' does not have an Excel extension (*.xls, *.xlsx)";

    my $xls_format = $xls_ext eq 'xls' ? xlNormal : $OpenXMLWorkbook;

    my $xls_abs = File::Spec->rel2abs($xls_name); $xls_abs =~ s{/}'\\'xmsg;

    my $ole_excel = get_excel() or croak "Can't start Excel";
    my $xls_book  = $ole_excel->Workbooks->Add or croak "Can't Workbooks->Add xls_abs '$xls_abs'";
    my $xls_sheet = $xls_book->Worksheets(1) or croak "Can't find Sheet '1' in xls_abs '$xls_abs'";

    $xls_book->SaveAs($xls_abs, $xls_format);
    $xls_book->Close;
}

sub tmp_book {
    my $ole_excel = get_excel() or croak "Can't start Excel";
    my $xls_book  = $ole_excel->Workbooks->Add or croak "Can't Workbooks->Add";

    return $xls_book;
}

sub get_excel {
    return $ole_global if $ole_global;

    # use existing instance if Excel is already running
    my $ol1 = eval { Win32::OLE->GetActiveObject('Excel.Application') };
    return if $@;

    unless (defined $ol1) {
        $ol1 = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
          or return;
    }
 
    $ole_global = $ol1;
    $ole_global->{DisplayAlerts} = 0;

    return $ole_global;
}

sub get_book {
    my ($prm_book_name) = @_;
 
    unless ($prm_book_name =~ m{\. xls x? \z}xmsi) {
        croak "xls_name '$prm_book_name' does not have an Excel extension (*.xls, *.xlsx)";
    }
 
    unless (-f $prm_book_name) {
        croak "xls_name '$prm_book_name' not found";
    }
 
    my $prm_book_abs = File::Spec->rel2abs($prm_book_name); $prm_book_abs =~ s{/}'\\'xmsg;
 
    my $obj_excel = get_excel()                                or croak "Can't start Excel";
    my $obj_book  = $obj_excel->Workbooks->Open($prm_book_abs) or croak "Can't Workbooks->Open xls_abs '$prm_book_abs'";
 
    return $obj_book;
}

sub get_last_row {
   $_[0]->UsedRange->Find({
     What            => '*',
     SearchDirection => xlPrevious,
     SearchOrder     => xlByRows,
   })->{Row};
}

sub get_last_col {
   $_[0]->UsedRange->Find({
     What            => '*',
     SearchDirection => xlPrevious,
     SearchOrder     => xlByColumns,
   })->{Column};
}

1;

__END__

=head1 NAME

Win32::Scsv - Convert from and to *.xls, *.csv using Win32::OLE

=head1 SYNOPSIS

    use Win32::Scsv qw(
      xls_2_csv csv_2_xls xls_2_vbs slurp_vbs empty_xls get_xver
      open_xls_book open_xls_sheet get_last_row get_last_col open_excel
    );

    my ($ver, $product) = get_xver;

    xls_2_csv('Test1.xls');
    xls_2_csv('Test1.xls' => 'dummy.csv');
    csv_2_xls('dummy.csv' => 'Test2.xls');
    xls_2_vbs('Test1.xls' => 'dummy.vbs');
    empty_xls('Test2.xls');

    say $_->{'NAME'}, ' => ', $_->{'CODE'} for @{slurp_vbs('Test3.xls')};

    csv_2_xls('dummy.csv' => 'New.xlsx%Tab9', {
      'tpl'  => 'Template.xls',
      'prot' => 1,
      'csz'  => [
         ['H:H' => 13.71],
         ['A:D' => 3],
      ],
      'fmt'  => [
         ['A:A' => '#,##0.000'],
         ['B:B' => '\\<@\\>'],
         ['C:C' => 'dd/mm/yyyy hh:mm:ss'],
      ],
    });

    my $ob = get_book('Test01.xls');
    my $os = $ob->Worksheets('Sheet5') or die "Can't find Sheet";

    my $last_row = get_last_row($os);
    my $last_col = get_last_col($os);

    say 'last row = ', $last_row, ', last col = ', $last_col;

    $ob->Close;

    open_excel('C:\Data\Test01.xls');

=head1 AUTHOR

Klaus Eichner <klaus03@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2009-2011 by Klaus Eichner

All rights reserved. This program is free software; you can redistribute
it and/or modify it under the terms of the artistic license 2.0,
see http://www.opensource.org/licenses/artistic-license-2.0.php

=cut
