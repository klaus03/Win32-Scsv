use 5.008;

use ExtUtils::MakeMaker;

use Win32::OLE::Const 'Microsoft Excel';
my $avail = xlCSV; # test that Excel is readily available

# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME              => 'Win32::Scsv',
    VERSION_FROM      => 'lib/Win32/Scsv.pm', # finds $VERSION
    PREREQ_PM         => { 'Win32::OLE'          => 0,
                           'Win32::OLE::Variant' => 0 }, # e.g., Module::Name => 1.1
    ($] >= 5.005 ?     ## Add these new keywords supported since 5.005
      (ABSTRACT_FROM  => 'lib/Win32/Scsv.pm', # retrieve abstract from module
       AUTHOR         => 'Klaus Eichner <klaus03@gmail.com>') : ()),
    ($ExtUtils::MakeMaker::VERSION >= 6.3002 ? ('LICENSE'=> 'perl') : ()),
    (eval { ExtUtils::MakeMaker->VERSION(6.46) } ? (META_MERGE => {
        'meta-spec' => { version => 2 },
        resources => {
            repository => {
                type => 'git',
                url  => 'https://github.com/klaus03/Win32-Scsv.git',
                web  => 'https://github.com/klaus03/Win32-Scsv',
            }
        }})
        : ()
    ),
);
