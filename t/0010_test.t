use strict;
use warnings;

use Test::More tests => 4;

use_ok('Win32::Scsv', qw(XLRef));

is(XLRef(1,  1),   'A1',   'XLRef Test 01');
is(XLRef(28, 22),  'AB22', 'XLRef Test 02');
is(XLRef(2,  456), 'B456', 'XLRef Test 03');
