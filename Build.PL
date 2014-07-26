use strict;
use warnings;
use 5.008;

use Module::Build;

use Win32::OLE::Const 'Microsoft Excel';
my $avail = xlCSV; # test that Excel is readily available

Module::Build->new(
  module_name        => 'Win32::Scsv',
  license            => 'perl',
  configure_requires => { 'Module::Build' => 0 },
  requires           => { 'Win32::OLE' => 0, 'Win32::OLE::Variant' => 0 },
  dist_abstract      => 'Convert from and to *.xls, *.csv using Win32::OLE',
)->create_build_script;
