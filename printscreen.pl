#!perl -w
#
use strict;
use warnings;

use Win32::Console::ANSI;
use Term::ANSIScreen qw/:color :cursor :screen/;
 
locate 1, 1; print "@ This is (1,1)", savepos;
print locate(24,60), "@ This is (24,60)"; loadpos;
print down(2), clline, "@ This is (3,16)\n";
color 'black on white'; clline;
print "This line is black on white.\n";
print color 'reset'; print "This text is normal.\n";
print colored ("This text is bold blue.\n", 'bold blue');
print "This text is normal.\n";
print colored ['bold blue'], "This text is bold blue.\n";
print "This text is normal.\n";
