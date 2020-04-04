#!perl -w
#

    use strict;
    # use Spreadsheet::ParseExcel;
    use DateTime::Format::ISO8601;
    use Spreadsheet::ParseXLSX;

    my $FileName = "/cygwin64/home/tzysnv/Report.xlsx";
    # my $FileName = "/cygwin64/home/tzysnv/Report.xls";
    my $parser = Spreadsheet::ParseXLSX->new;
    # my $parser   = Spreadsheet::ParseExcel->new();
    my $workbook = $parser->parse($FileName);

    die $parser->error(), ".\n" if ( !defined $workbook );

    # Following block is used to Iterate through all worksheets
    # in the workbook and print the worksheet content 

        my $worksheet = $workbook->worksheet('Sheet1');
	my $name = $worksheet->get_name();

        # Find out the worksheet ranges
        my ( $row_min, $row_max ) = $worksheet->row_range();
        my ( $col_min, $col_max ) = $worksheet->col_range();

	#print "worksheet_name = $name\n";
	#print "row_min, row_max = ($row_min, $row_max)\n";
	#print "col_min, col_max = ($col_min, $col_max)\n";

	$row_min = 1;
	my $col_sw_pr_num = 0;
	my $col_status = 2;
	my $col_email_mod = 14;
	my $col_create_date = 27;
	my $col_complete_date = 32;
	my $strOffset = 0;
	my $date_pos = 10;

	my $count = 0; # total nr of SW Peer Reviews
	my $count19 = 0; # total nr of SW Peer Reviews in 2019
	my $count201 = 0; # total nr of SW Peer Reviews in Jan 2020
	my $count202 = 0; # total nr of SW Peer Reviews in Feb 2020
	my $count203 = 0; # total nr of SW Peer Reviews in Mar 2020
	my $openJan2020 = 0;
	my $openFeb2020 = 0;
	my $openMar2020 = 0;
	my $opened = 0;  # Opened nr of SW Peer Reviews
	my $email_start_pos = 20;
	my $all_emails = "abdul.majeed\@delphi.com;hadrian.ho\@delphi.com;ronnie.kartha\@delphi.com;wai.phyo.zaw\@delphi.com;ser.gin.chia\@delphi.com;kean.hin.lee\@delphi.com;goke.how.lee\@delphi.com;raymond.tan\@delphi.com;tse.tsong.teo\@delphi.com;feng.qi\@delphi.com;ronald.yong\@delphi.com;steven.liu\@delphi.com";
	my $dtoday = DateTime->today;

	#printf("%3s  | %15s | %20s | %15s | %30s | %30s | %12s |\n", "Nr","SW PR#","Status","Create Date","Author Email","Moderator Email","Elapsed Day");
        print "", pad('Nr', 5), "|";
        print "", pad('SW PR#', 17), "|";
        print "", pad('Status', 22), "|";
        print "", pad('Create Date', 17), "|";
        print "", pad('Author Email', 32), "|";
        print "", pad('Moderator Email', 32), "|";
        print "", pad('Elapsed Day', 14), "|\n";
	print "--------------------------------------------------------------------------------------------------------------------------------------------------\n";

	for my $row ( $row_min .. $row_max ) {

	    # get sw pr num
            my $sw_pr_num = $worksheet->get_cell( $row, $col_sw_pr_num )->value();
	    # get status
            my $sw_status = $worksheet->get_cell( $row, $col_status )->value();
	    # get moderator email
	    my $mod_email = $worksheet->get_cell( $row, $col_email_mod )->value();
            # create date of SW Peer Review
            my $create_date = $worksheet->get_cell( $row, $col_create_date );
	    my $str1 = $create_date->value();
	    my $length1 = length $str1;
	    my $date1 = substr($str1,$strOffset,$date_pos);
	    my $email1 = substr($str1,$email_start_pos,$length1-$email_start_pos);
	    my $idx1 = index($all_emails,$email1);

            # completed date of SW Peer Review
	    my $idx2 = -1;
	    my $date2 = "";
	    my $email2 = "";
            my $completed_date = $worksheet->get_cell( $row, $col_complete_date );
	    
	    if (defined $completed_date) {
	    	my $str2 = $completed_date->value();
	        my $length2 = length $str2;
	        $date2 = substr($str2,$strOffset,$date_pos);
	        $email2 = substr($str2,$email_start_pos,$length2-$email_start_pos);
	        $idx2 = index($all_emails,$email2);
	    } 

	    #print "email1=\"$email1\" email2=\"$email2\"\n";
	    #print "idx1=$idx1 idx2=$idx2\n";

	    if (($idx1>=0) || ($idx2>=0)) {
		# print "($row. ", $create_date->value(), ",", $completed_date->value(), ")\n";
		if ($date2 eq "") {
		    # Opened SW Peer Reviews
		    $opened++;
		    my $dt = DateTime::Format::ISO8601->parse_datetime( $date1 );
		    my $diff = $dtoday->delta_days($dt)->delta_days;
		    #print "$opened. $sw_pr_num $sw_status $date1 $email1 Delta days = $diff\n";
		    #printf("%3s. | %15s | %20s | %15s | %30s | %30s | %7s      | \n",$opened,$sw_pr_num,$sw_status,$date1,$email1,$mod_email,$diff);
        	    print "", pad($opened, 5), "|";
        	    print "", pad($sw_pr_num, 17), "|";
        	    print "", pad($sw_status, 22), "|";
        	    print "", pad($date1, 17), "|";
        	    print "", pad($email1, 32), "|";
        	    print "", pad($mod_email, 32), "|";
        	    print "", pad($diff, 14), "|\n";
		}
		else { # Calculate overdue SW Peer Reviews that have been closed
		    my $dt1 = DateTime::Format::ISO8601->parse_datetime( $date1 );
		    my $dt2 = DateTime::Format::ISO8601->parse_datetime( $date2 );
		    my $diff = $dt2->delta_days($dt1)->delta_days;
		    if ($diff > 60) {
			    # print "Overdue : $diff days, $date1 $email1 $date2 $email2\n";
		    } 
		}

		if (substr($date1,$strOffset,4) eq "2019") { # count all 2019 items
		    $count19++;
		    if ($date2 eq "") { # created in 2019 but still not closed
			$openJan2020++;
			$openFeb2020++;
			$openMar2020++;
		    }
		    elsif (substr($date2,$strOffset,7) eq "2020-02") {
			$openJan2020++;
		    } 
		    elsif (substr($date2,$strOffset,7) eq "2020-03") {
			$openJan2020++;
			$openFeb2020++;
		    } 
		}
		else {
		    if (substr($date1,$strOffset,7) eq "2020-01") { # count all 2020 Jan items
		        $count201++;
		        if ($date2 eq "") { # created in Jan 2020 but still not closed
			    $openJan2020++;
			    $openFeb2020++;
			    $openMar2020++;
		        }
		        elsif (substr($date2,$strOffset,7) eq "2020-02") { # created in Jan 2020 and closed in Feb 2020
			    $openJan2020++;
		        } 
		        elsif (substr($date2,$strOffset,7) eq "2020-03") { # created in Jan 2020 and closed in Mar 2020
			    $openJan2020++;
			    $openFeb2020++;
		        } 
		    }
		    elsif (substr($date1,$strOffset,7) eq "2020-02") { # count all 2020 Feb items
		        $count202++;
		        if ($date2 eq "") { # created in Feb 2020 but still not closed
			    $openFeb2020++;
			    $openMar2020++;
		        }
		        elsif (substr($date2,$strOffset,7) eq "2020-03") { # created in Feb 2020 and closed in Mar 2020
			    $openFeb2020++;
		        } 
		    }
		    elsif (substr($date1,$strOffset,7) eq "2020-03") { # count all 2020 Mar items
		        $count203++;
		        if ($date2 eq "") { # created in Mar 2020 but still not closed
			    $openMar2020++;
		        }
		    }
		}

		$count++;
		next;
	    } 
        }
	print "Total nr of SW Peer Reviews = $count \n";
	print "Total nr of SW Peer Reviews in 2019 = $count19 \n";
	print "Total nr of SW Peer Reviews in Jan 2020 = ",$count19+$count201,"\n";
	print "Total nr of Open SW Peer Reviews in Jan 2020 = $openJan2020 \n";
	print "Total nr of SW Peer Reviews in Feb 2020 = ",$count19+$count201+$count202,"\n";
	print "Total nr of Open SW Peer Reviews in Feb 2020 = $openFeb2020 \n";
	print "Total nr of SW Peer Reviews in Feb 2020 = ",$count19+$count201+$count202+$count203,"\n";
	print "Total nr of Open SW Peer Reviews in Mar 2020 = $openMar2020 \n";
	print "Total nr of Opened SW Peer Reviews = $opened \n";


sub pad {
    # Return $str centered in a field of $col $padchars.
    # $padchar defaults to ' ' if not specified.
    # $str is truncated to len $column if too long.

    my ($str, $col, $padchar) = @_;
    $padchar = ' ' unless $padchar;
    my $strlen = length($str);
    $str = substr($str, 0, $col) if ($strlen > $col);
    my $fore = int(($col - $strlen) / 2);
    my $aft = $col - ($strlen + $fore);
    $padchar x $fore . $str . $padchar x $aft;
}

