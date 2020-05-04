#!perl -w
use strict;
use DateTime;
use DateTime::Format::Strptime qw( );
my $dt = DateTime->now(time_zone=>'local');
print "Date1 is : ",$dt->strftime("%d-%b-%Y"),"\n";
print "Date2 is : ",$dt->strftime("%d-%m-%Y"),"\n";
print "Time1 is : ",$dt->strftime("%H:%M:%S %Z"),"\n";
print "Time2 is : ",$dt->strftime("%I:%M:%S %p %Z"),"\n";

my $dt3 = DateTime->now->subtract(days => 1);
print "This is my date:$dt3\n";

print "This is my date: ", $dt3->ymd(''), "\n";

print "This is my date: ", $dt3->strftime('%Y%m%d'), "\n";

my $format = DateTime::Format::Strptime->new( pattern => '%Y%m%d' );
print "This is my date: ", $format->format_datetime($dt3), "\n";

my $dt3 = DateTime->now(time_zone=>'local')->subtract(days => 1);
print "This is my date:$dt3\n";

print "This is my date: ", $dt3->ymd(''), "\n";

print "This is my date: ", $dt3->strftime('%Y%m%d'), "\n";

my $format = DateTime::Format::Strptime->new( pattern => '%Y%m%d' );
print "This is my date: ", $format->format_datetime($dt3), "\n";
