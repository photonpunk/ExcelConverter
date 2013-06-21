##  Copyright (c) 2013, The University of Alabama Libraries.
##  Contributed by {Austin Dixon}  {5/13/2013}.
##  All rights reserved.
 
##  Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
 
##    * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
##    * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the
##       distribution.
##    * Neither the name of The University of Alabama Libraries nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
 
##THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
##THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
##CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
##PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
##LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE,
##EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.




# script creates either an tab-delimited, unicode tab-delimited, csv, or xml file from a user selected excel spreadsheet
# Unicode option also parses and removes only the unwanted quotes added by the Excel export. the *good* quotes within a cell are retained
# Unicode option now re-encodes the exported UCS-2 16bit little endian into standard UTF-8 without BOM
 
 
#! usr/local/bin/perl
use Tkx;
use LWP::Simple;
use File::Copy;
use Win32::OLE qw(in with);
BEGIN { Win32::OLE->Initialize( Win32::OLE::COINIT_OLEINITIALIZE() ) }
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);
 
# SETUP
    											
# frame objects
my $fa;
my $fb;
my $fc;
# label objects
my $la;
my $lb;
my $lc;
my $ld;
my $le;
# button objects
my $ba;
my $bb;
my $bc;
my $bd;
my $bd;
my $be;
my $bf;
 
my $new_url = ''; # default url to save new file to
 
my $choice = 1; # default filetype (unicode)
 
#-------------------------------------------------
# convert button sub routine
 
sub chooseFile {
$url = Tkx::tk___getOpenFile();
$bc->configure( -state => 'active');
} 
 
#-------------------------------------------------
# file convert sub routine
 
sub runFileConverter {
 
if (substr($new_url, 2, 0) == '/') # determines if user selected address uses forward or backward slashes
{
	$s = '/';
}
else
{
	$s = '\\';
}
 
$Win32::OLE::Warn = 3; # Die on Errors.
 
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
            || Win32::OLE->new('Excel.Application', 'Quit');
 
$Excel->{DisplayAlerts}=0;
 
my $excel_file = $url;
my $workbook = $Excel->Workbooks->Open($excel_file);
 
 
if ($choice == 1) # tab-delimited
{
	my $saveas_file = 'C:\ExcelConverter\OUTPUT\tab.txt';
 
	$workbook->SaveAs($saveas_file, -4158);				# see XlFileFormat Enumeration for more info
	sleep(4);
	#-------
	my $t = 'tab.txt';
	my $old_file = $saveas_file;
	my $new_file = $new_url . $s . $t;  				# builds new directory path
	
	copy($old_file,$new_file) or die "Copy failed: $!";	# renames (moves) files
}
 
elsif ($choice == 2) # csv
{
	my $saveas_file = 'C:\ExcelConverter\OUTPUT\csv.txt';
 
	$workbook->SaveAs($saveas_file, 6);					# see XlFileFormat Enumeration for more info
	sleep(4);
	#-------
	my $c = 'csv.txt';
	my $old_file = $saveas_file;
	my $new_file = $new_url . $s . $c; 	  				# builds new directory path
 
	copy($old_file,$new_file) or die "Copy failed: $!";	# renames (moves) files
}
 
elsif ($choice == 3) # xml
{
	my $saveas_file = 'C:\ExcelConverter\OUTPUT\xml.xml';
 
	$workbook->SaveAs($saveas_file, 46);				# see XlFileFormat Enumeration for more info
	sleep(4);
	#-------
	my $x = 'xml.xml';
	my $old_file = $saveas_file;
	my $new_file = $new_url . $s . $x;		  			# builds new directory path
 
	copy($old_file,$new_file) or die "Copy failed: $!";	# renames (moves) files
 
}
 
elsif ($choice == 4) 												# unicode (UCS-2 Little-Endian) (tab-delimited)
{
	my $saveas_file = 'C:\ExcelConverter\OUTPUT\uni.txt';
 
	$workbook->SaveAs($saveas_file, 42);							# see XlFileFormat Enumeration for more info
	sleep(4);														# pause for file saving
	#-------
	
	open(my $rawbatchexport, "<", $saveas_file) or die; 			# opens the excel unicode text export file (in read mode) so we can clean up all that dirty quoted filth
	binmode($rawbatchexport, ":encoding(UTF-16LE)");				# change binary encoding mode to reflect the files UCS-2 (UTF-16LE) encoding
	
	my $u = 'cleaned_UTF8.txt';										# new filename and extention
	my $new_file = $new_url . $s . $u;					  			# builds new directory path
 
	open(my $cleanbatchfile, ">>", $new_file) or die; 				# open a new file (in append mode) for printing the cleaned up lines into
	binmode($cleanbatchfile, ":encoding(UTF-8)");					# change binary encoding mode of the output file to UTF-8 (which by default is "without BOM")
	
	while ($line = <$rawbatchexport>) {								# while loop that read each line of a file
		chomp($line);
		$line =~ s/(?!"")"//g; 										# a substitution of nil for characters that satisfy this match will remove all stand alone double-quotes and reduce sequences of 3 to 2, and sequences of 2 to 1.
		$line =~ s/""/"/g; 											# a follow up substitution regex will replace all remaining sequences of 2 double-quotes with 1
		print $cleanbatchfile $line;								# write the cleaned line of metadata to the new file
	}
	
	close $rawbatchexport or die;									# close dirty excel unicode export
	unlink $saveas_file;											# delete old export
	close $cleanbatchfile or die;									# close cleaned up file
 
}
}
 
#-------------------------------------------------
# GUI elements
 
# main window
my $mw = Tkx::widget->new(".");
$mw->g_wm_title("Excel Converter");
$mw->g_wm_minsize(150, 130);
 
# frame a
$fa = $mw->new_frame(
-relief => 'solid',
-borderwidth => 1,
-background => 'light gray',
);
$fa->g_pack( -side => 'left', -fill => 'both' );
 
#---------------------------------------------------
# choose file
 
$la = $fa->new_label(
-text => 'Choose File to Scan:',
-font => 'bold',
-bg => 'light gray',
-foreground => 'black',
);
$la->g_pack( -side => 'top', -fill => 'both' );
 
$lb = $fa->new_label(
-bg => 'blue',
-foreground => 'cyan',
-width => 28,
-textvariable => \$url,
);
$lb->g_pack( -side => 'top' );
 
$ba = $fa->new_button(
-text => 'Choose',
-command => \&chooseFile,
-height => 1,
-width => 15,
);
$ba->g_pack( -side => 'top', -pady => 5 );
 
#---------------------------------------------------
# choose directory
 
$ld = $fa->new_label(
-text => '   Choose Where To Save File:   ',
-font => 'bold',
-bg => 'light gray',
);
$ld->g_pack( -side => 'top', -fill => 'both' );
 
$le = $fa->new_label(
-bg => 'blue',
-foreground => 'cyan',
-width => 28,
-textvariable => \$new_url,
);
$le->g_pack( -side => 'top' );
 
$bf = $fa->new_button(
-text => 'Choose',
-command => sub {$new_url = Tkx::tk___chooseDirectory();},
-height => 1,
-width => 15,
);
$bf->g_pack( -side => 'top', -pady => 5 );
 
#------------------------------------------------
# convert button
 
$bc = $fa->new_button(
-borderwidth => 1,
-text => 'Convert File!',
-font => 'bold',
-command => \&runFileConverter,
-state => 'disabled',
-height => 2,
-width => 15,
);
$bc->g_pack( -side => 'bottom', -pady => 10 );
 
#------------------------------------------------
# frame b (choose filetype)
 
$fb = $mw->new_frame(
-relief => 'solid',
-borderwidth => 1,
-background => 'light gray',
);
$fb->g_pack( -side => 'right', -fill => 'both' );
 
$lb = $fb->new_label(
-text => 'Choose Filetype:',
-font => 'bold',
-bg => 'light gray',
-foreground => 'black',
);
$lb->g_pack( -side => 'top', -fill => 'both' );
 
$bb = $fb->new_radiobutton(
-bg => 'light gray',
-text => "Tab-Delimited", 
-variable => \$choice, 
-value => 1,	
);
$bb->g_pack( -anchor => w, -side => 'top', -pady => 4 );
 
$bd = $fb->new_radiobutton(
-bg => 'light gray',
-text => "CSV", 
-variable => \$choice, 
-value => 2,	
);
$bd->g_pack(  -anchor => w, -side => 'top', -pady => 4);
 
$be = $fb->new_radiobutton(
-bg => 'light gray',
-text => "XML", 
-variable => \$choice, 
-value => 3,	
);
$be->g_pack(  -anchor => w, -side => 'top', -pady => 4 );
 
$bf = $fb->new_radiobutton(
-bg => 'light gray',
-text => "Unicode", 
-variable => \$choice, 
-value => 4,	
);
$bf->g_pack(  -anchor => w, -side => 'top', -pady => 4 );
 
Tkx::MainLoop();
