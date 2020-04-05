#!perl -w
#
# changes : format output

    use strict;
    # use Spreadsheet::ParseExcel;
    use DateTime::Format::ISO8601;
    use Spreadsheet::ParseXLSX;
    use List::Util qw(min max);

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
	# initialize arrays
	my @nr;
	my @arr_sw_pr_num;
	my @arr_sw_status;
	my @arr_date1;
	my @arr_email1;
	my @arr_mod_email;
	my @arr_diff;

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
		    my $dt = DateTime::Format::ISO8601->parse_datetime( $date1 );
		    my $diff = $dtoday->delta_days($dt)->delta_days;
		    $nr[$opened] = $opened + 1;
        	    $arr_sw_pr_num[$opened] = $sw_pr_num;
        	    $arr_sw_status[$opened] = $sw_status;
        	    $arr_date1[$opened] = $date1;
        	    $arr_email1[$opened] = $email1;
        	    $arr_mod_email[$opened] = $mod_email;
        	    $arr_diff[$opened] = $diff;
		    $opened++;
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

	my $nrWidth = max (map length, @nr, 'Nr') + 3;
	my $prnumWidth = max (map length, @arr_sw_pr_num, 'SW PR#') + 3;
	my $statusWidth = max (map length, @arr_sw_status, 'Status') + 3;
	my $date1Width = max (map length, @arr_date1, 'Create Date') + 3;
	my $email1Width = max (map length, @arr_email1, 'Author Email') + 3;
	my $modemailWidth = max (map length, @arr_mod_email, 'Moderator Email') + 3;
	my $diffWidth = max (map length, @arr_diff, 'Elapsed Day') + 3;

	print '-' x ($nrWidth + $prnumWidth + $statusWidth + $date1Width + $email1Width + $modemailWidth + $diffWidth), "\n";
	printf "%-*s%*s%*s%*s%*s%*s%*s\n",
    		$nrWidth, "Nr", $prnumWidth, "SW PR#", $statusWidth, "Status", $date1Width, "Create Date",
    		$email1Width, "Author Email", $modemailWidth, "Moderator Email", $diffWidth, "Elapsed Day";
	print '-' x ($nrWidth + $prnumWidth + $statusWidth + $date1Width + $email1Width + $modemailWidth + $diffWidth), "\n";

	for my $index (0 .. $#nr) {
		printf "%-*s%*s%*s%*s%*s%*s%*s\n",
    		$nrWidth, $nr[$index], $prnumWidth, $arr_sw_pr_num[$index],
    		$statusWidth, $arr_sw_status[$index], $date1Width, $arr_date1[$index],
    		$email1Width, $arr_email1[$index], $modemailWidth, $arr_mod_email[$index], $diffWidth, $arr_diff[$index];
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

