#!perl -w
# this code doesnt work cos Mail Outlook not supported

use strict;
use warnings;


# create the object
use Mail::Outlook;
my $outlook = new Mail::Outlook();
 
# start with a folder
my $outlook = new Mail::Outlook('Inbox');
 
# use the Win32::OLE::Const definitions
use Mail::Outlook;
use Win32::OLE::Const 'Microsoft Outlook';
my $outlook = new Mail::Outlook(olInbox);
 
# get/set the current folder
my $folder = $outlook->folder();
my $folder = $outlook->folder('Inbox');
 
# get the first/last/next/previous message
my $message = $folder->first();
   $message = $folder->next();
   $message = $folder->last();
   $message = $folder->previous();
 
# read the attributes of the current message
my $text = $message->From();
   $text = $message->To();
   $text = $message->Cc();
   $text = $message->Bcc();
   $text = $message->Subject();
   $text = $message->Body();
my @list = $message->Attach();
 
# use Outlook to display the current message
$message->display;
 
# create a message for sending
my $message = $outlook->create();
$message->To('vincent.chan@delphi.com');
#$message->Cc('Them <them@example.com>');
#$message->Bcc('Us <us@example.com>; anybody@example.com');
$message->Subject('Blah Blah Blah');
$message->Body('Yadda Yadda Yadda');
#$message->Attach(@lots_of_files);
#$message->Attach(@more_files);    # attachments are appended
#$message->Attach($one_file);      # so multiple calls are allowed
$message->send;
 
# Or use a hash
my %hash = (
   To      => 'vincent.chan@delphi.com',
#  Cc      => 'Them <them@example.com>',
#   Bcc     => 'Us <us@example.com>, anybody@example.com',
   Subject => 'Hash Blah Blah Blah',
   Body    => 'Yadda Yadda Yadda',
);
 
my $message = $outlook->create(%hash);
$message->display(%hash);
$message->send(%hash);
