use strict;
use warnings;

use Test::More tests => 12;

use_ok('Win32::Scsv', qw(XLRef XLConst));

is(XLRef(1,  1),   'A1',   'XLRef Test 01');
is(XLRef(28, 22),  'AB22', 'XLRef Test 02');
is(XLRef(2,  456), 'B456', 'XLRef Test 03');
is(XLRef(30),      'AD',   'XLRef Test 04');

my $CN = XLConst();

is ($CN->{'xlNormal'},             -4143, 'Test xlNormal');
is ($CN->{'xlPasteValues'},        -4163, 'Test xlPasteValues');
is ($CN->{'xlCSV'},                    6, 'Test xlCSV');
is ($CN->{'xlCalculationManual'},  -4135, 'Test xlCalculationManual');
is ($CN->{'xlPrevious'},               2, 'Test xlPrevious');
is ($CN->{'xlByRows'},                 1, 'Test xlByRows');
is ($CN->{'xlByColumns'},              2, 'Test xlByColumns');
