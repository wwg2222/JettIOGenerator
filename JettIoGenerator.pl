# Scan IODefine.cpp to generate _io_defs.h and _io_defs.cpp
#
# $Version should be changed everytime this file is edited
# Depending on amount of changes, add 0.01 or 0.1 as you see fit 
$version = "1.01";

use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use File::Basename;

$Win32::OLE::Warn = 3;

if ($#ARGV < 0) {
  print "Usage ".basename($0)." FullPath2IOList.xls";
  #exit(-1);
  
  
  $ARGV[0] = 'E:\cvslocal\Projects\Tools\IOExcelReader\IE LP IO List  ZQZ 052010.xls'
}

my $excel_file = $ARGV[0];
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')  || Win32::OLE->new('Excel.Application', 'Quit');  
my $Book = $Excel->Workbooks->Open($excel_file); 


my $dest_file = ".\\IODefine.cpp";
my $temp_file = ".\\temp.txt";
open(OUT_FILE, ">$dest_file") || die "Could't open output file $dest_file";
open(DEBUG_FILE, ">$temp_file") || die "Could't open output file $temp_file";


my $DeviceID   = 0;

my $BLOCK_SYSTEM_AI = 0;
my $BLOCK_DI = 1;
my $BLOCK_AI = 2;
my $BLOCK_MOTION_STATUS = 3;   
my $BLOCK_TEMPER_STATUS = 4;
my $BLOCK_GASBOX_STATUS = 5;

my $BLOCK_SYSTEM_AO = 6;
my $BLOCK_DO = 7;
my $BLOCK_AO = 8;
my $BLOCK_MOTION_CONTROL = 9; 
my $BLOCK_TEMPER_CONTROL = 10;
my $BLOCK_GASBOX_CONTROL = 11;


my $BLOCK_SAFETY_IN     = $BLOCK_MOTION_STATUS;   #Safty/motion share block
my $BLOCK_SAFETY_OUT    = $BLOCK_MOTION_CONTROL;   #Safty/motion share block 

my $IO_FLAG_ASCII       = 32; 
my $IO_FLAG_BIPOLAR     = 8; 

my $StartRow = 3;
my $Sheet_IOList  = 3;
my $Sheet_ProList = 4;
my $Sheet_SystemMapping = 5;

my $ITEM_TYPE     = 0;
my $ITEM_COL      = 1;
my $ITEM_BLOCK    = 2;
my $ITEM_INDEX    = 3;
my $ITEM_SUBTYPE  = 4;

my $diCount = 0;
my $doCount = 0;
my $aiCount = 0;
my $aoCount = 0;

my $PLC_BITS = 16;

&CreateSystemAI();
printf OUT_FILE  "\r\n\r\n";

&CreateDirectIn();
printf OUT_FILE  "\r\n\r\n";

&CreateDirectAI();
printf OUT_FILE  "\r\n\r\n";

&CreateSaftyIn();
printf OUT_FILE  "\r\n\r\n";

&CreateProfibusIn();
printf OUT_FILE  "\r\n\r\n";

&CreateGasBoxIn();
printf OUT_FILE  "\r\n\r\n";

&CreateSystemAO();
printf OUT_FILE  "\r\n\r\n";

&CreateDirectOut();
printf OUT_FILE  "\r\n\r\n";

&CreateSaftyOut();
printf OUT_FILE  "\r\n\r\n";

&CreateProfibusOut();
printf OUT_FILE  "\r\n\r\n";

&CreateGasBoxOut();

$Book->Close;


close(OUT_FILE);
close(DEBUG_FILE);


sub CreateSystemAI
{
my $Sheet = $Book->Worksheets($Sheet_SystemMapping);
my @Items =(                          
                ["SystemAI",  "C",  "$BLOCK_SYSTEM_AI", "0",   "0"], 			 
             );  
               

	my $ioOffset = 0;
	my $bitOffset = 0;
    for my $i(0..$#Items)
    {  

        my $Type    = $Items[$i][$ITEM_TYPE];
        my $Col     = $Items[$i][$ITEM_COL];
        my $Block   = $Items[$i][$ITEM_BLOCK];
        my $Index   = $Items[$i][$ITEM_INDEX];
        my $SubType = $Items[$i][$ITEM_SUBTYPE];
		
#		printf OUT_FILE  "\r\n\/\/%s define\r\n", $Type;
			
        if( length($Col) >1)
        {
          my @Chars = split "",$Col;
          $Col = ord($Chars[1]) - ord('A') + 26 + 1;
        }
        else
        {
         $Col = ord($Col) - ord('A') + 1;
        } 
 
        for(my $Row = 4; $Row <= 45; $Row++)
        {  
         my $Name = $Sheet->Cells($Row,$Col)->{'Value'};
         printf  DEBUG_FILE "%s %s %s %s %s\n",$Name, $Type, $Block , $Index, $SubType;
         my $Bits = &CreateIO($Name, $Type, $Block , $Index, $SubType, $ioOffset, $bitOffset);
      
        $Index++;
		$bitOffset += $Bits;
		if ($bitOffset >= $PLC_BITS) {
			$ioOffset += $bitOffset / $PLC_BITS;
			$bitOffset %= $PLC_BITS;
		}
       }
    
       print  OUT_FILE "\n\n"; 
    }
}

sub CreateSystemAO
{
my $Sheet = $Book->Worksheets($Sheet_SystemMapping);
my @Items =(                          
                ["SystemAO",  "F",  "$BLOCK_SYSTEM_AO", "0",   "0"],		 
             );  
               

	my $ioOffset = 0;
	my $bitOffset = 0;
    for my $i(0..$#Items)
    {  

        my $Type    = $Items[$i][$ITEM_TYPE];
        my $Col     = $Items[$i][$ITEM_COL];
        my $Block   = $Items[$i][$ITEM_BLOCK];
        my $Index   = $Items[$i][$ITEM_INDEX];
        my $SubType = $Items[$i][$ITEM_SUBTYPE];
		
#		printf OUT_FILE  "\r\n\/\/%s define\r\n", $Type;
			
        if( length($Col) >1)
        {
          my @Chars = split "",$Col;
          $Col = ord($Chars[1]) - ord('A') + 26 + 1;
        }
        else
        {
         $Col = ord($Col) - ord('A') + 1;
        } 
 
        for(my $Row = 4; $Row <= 4; $Row++)
        {  
         my $Name = $Sheet->Cells($Row,$Col)->{'Value'};
         printf  DEBUG_FILE "%s %s %s %s %s\n",$Name, $Type, $Block , $Index, $SubType;
         my $Bits = &CreateIO($Name, $Type, $Block , $Index, $SubType, $ioOffset, $bitOffset);
      
        $Index++;
		$bitOffset += $Bits;
		if ($bitOffset >= $PLC_BITS) {
			$ioOffset += $bitOffset / $PLC_BITS;
			$bitOffset %= $PLC_BITS;
		}
       }
    
       print  OUT_FILE "\n\n"; 
    }
}

sub CreateDirectIn
{
my $Sheet = $Book->Worksheets($Sheet_IOList);
my @Items =( 
                ["LocalDI",  "B",   "$BLOCK_DI", "0",   "0"],
                ["PM1DI",    "X",   "$BLOCK_DI", "352", "0"],
                ["PM2DI",    "AF",  "$BLOCK_DI", "432", "0"],
                ["PM3DI",    "AF",  "$BLOCK_DI", "512", "0"],
                ["PM4DI",    "AF",  "$BLOCK_DI", "592", "0"],
               );  
               
    for my $i(0..$#Items)
    {  

        my $Type    = $Items[$i][$ITEM_TYPE];
        my $Col     = $Items[$i][$ITEM_COL];
        my $Block   = $Items[$i][$ITEM_BLOCK];
        my $Index   = $Items[$i][$ITEM_INDEX];
        my $SubType = $Items[$i][$ITEM_SUBTYPE];
		my $ioOffset = $Index / 16;
		my $bitOffset = $Index % 16;
		
#		printf OUT_FILE  "\r\n\/\/%s define\r\n", $Type;
			
        if( length($Col) >1)
        {
          my @Chars = split "",$Col;
          $Col = ord($Chars[1]) - ord('A') + 26 + 1;
        }
        else
        {
         $Col = ord($Col) - ord('A') + 1;
        } 
 
        foreach my $Row ($StartRow..$Sheet->{UsedRange}->{Rows}->{Count})
        {  
         my $Name = $Sheet->Cells($Row,$Col)->{'Value'};
         printf  DEBUG_FILE "%s %s %s %s %s\n",$Name, $Type, $Block , $Index, $SubType;
         my $Bits = CreateIO($Name, $Type, $Block , $Index, $SubType, $ioOffset, $bitOffset);
      
        $Index++;
		$bitOffset += $Bits;
		if ($bitOffset >= $PLC_BITS) {
			$ioOffset += $bitOffset / $PLC_BITS;
			$bitOffset %= $PLC_BITS;
		}
       }
    
       print  OUT_FILE "\n\n"; 
    }
}

sub CreateDirectOut
{
my $Sheet = $Book->Worksheets($Sheet_IOList);
my @Items =(                          
                ["LocalDO",  "F",  "$BLOCK_DO", "0",   "0"],
                ["PM1DO",    "AB", "$BLOCK_DO", "128", "0"],
                ["PM2DO",    "AJ", "$BLOCK_DO", "192", "0"],
                ["PM3DO",    "AJ", "$BLOCK_DO", "240", "0"],
                ["PM4DO",    "AJ", "$BLOCK_DO", "288", "0"], 
              );  
               

    for my $i(0..$#Items)
    {  

        my $Type    = $Items[$i][$ITEM_TYPE];
        my $Col     = $Items[$i][$ITEM_COL];
        my $Block   = $Items[$i][$ITEM_BLOCK];
        my $Index   = $Items[$i][$ITEM_INDEX];
        my $SubType = $Items[$i][$ITEM_SUBTYPE];
		my $ioOffset = $Index / 16;
		my $bitOffset = $Index % 16;
		
#		printf OUT_FILE  "\r\n\/\/%s define\r\n", $Type;
			
        if( length($Col) >1)
        {
          my @Chars = split "",$Col;
          $Col = ord($Chars[1]) - ord('A') + 26 + 1;
        }
        else
        {
         $Col = ord($Col) - ord('A') + 1;
        } 
 
        foreach my $Row ($StartRow..$Sheet->{UsedRange}->{Rows}->{Count})
        {  
         my $Name = $Sheet->Cells($Row,$Col)->{'Value'};
         printf  DEBUG_FILE "%s %s %s %s %s\n",$Name, $Type, $Block , $Index, $SubType;
         my $Bits = &CreateIO($Name, $Type, $Block , $Index, $SubType, $ioOffset, $bitOffset);
      
        $Index++;
		$bitOffset += $Bits;
		if ($bitOffset >= $PLC_BITS) {
			$ioOffset += $bitOffset / $PLC_BITS;
			$bitOffset %= $PLC_BITS;
		}
       }
    
       print  OUT_FILE "\n\n"; 
    }
}

sub CreateDirectAI
{
my $Sheet = $Book->Worksheets($Sheet_IOList);
my @Items =(                          
                ["LocalAI",  "J",  "$BLOCK_AI", "0",   "0"], 			 
             );  
               

	my $ioOffset = 0;
	my $bitOffset = 0;
    for my $i(0..$#Items)
    {  

        my $Type    = $Items[$i][$ITEM_TYPE];
        my $Col     = $Items[$i][$ITEM_COL];
        my $Block   = $Items[$i][$ITEM_BLOCK];
        my $Index   = $Items[$i][$ITEM_INDEX];
        my $SubType = $Items[$i][$ITEM_SUBTYPE];
		
#		printf OUT_FILE  "\r\n\/\/%s define\r\n", $Type;
			
        if( length($Col) >1)
        {
          my @Chars = split "",$Col;
          $Col = ord($Chars[1]) - ord('A') + 26 + 1;
        }
        else
        {
         $Col = ord($Col) - ord('A') + 1;
        } 
 
        foreach my $Row ($StartRow..$Sheet->{UsedRange}->{Rows}->{Count})
        {  
         my $Name = $Sheet->Cells($Row,$Col)->{'Value'};
         printf  DEBUG_FILE "%s %s %s %s %s\n",$Name, $Type, $Block , $Index, $SubType;
         my $Bits = &CreateIO($Name, $Type, $Block , $Index, $SubType, $ioOffset, $bitOffset);
      
        $Index++;
		$bitOffset += $Bits;
		if ($bitOffset >= $PLC_BITS) {
			$ioOffset += $bitOffset / $PLC_BITS;
			$bitOffset %= $PLC_BITS;
		}
       }
    
       print  OUT_FILE "\n\n"; 
    }
}

sub CreateSaftyIn
{
#     
#    printf OUT_FILE  "\r\n\/\/SafetyDI define\r\n";
    my $Sheet = $Book->Worksheets($Sheet_ProList);

#	my $index = 448;
	my $index = 0;
	my $flag  = 0;
	
	my $Bits = 1;
	my $ioOffset = 0;
	my $bitOffset = 0;
	for( my $Row = 4; $Row <= 19; $Row++)   #NotesI_1
	{
	    my $name = &GetName( $Sheet->Cells($Row,"K")->{'Value'}, $diCount + 1);
	    printf OUT_FILE  "DEFINE_PLC_DI\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DI\\\\CH%d\", \t\"\\\\PLC\\\\DI\\\\SAFTETY\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	        "di".$name.",", $DeviceID, $BLOCK_SAFETY_IN, $index, $flag,
	        $ioOffset, $bitOffset, $Bits, $diCount, $name, "Safety DI";

		$diCount++;
		$index++;
	    $bitOffset += $Bits;
	    if ($bitOffset >= $PLC_BITS) {
	    	$ioOffset += $bitOffset / $PLC_BITS;
	    	$bitOffset %= $PLC_BITS;
	    }
	}
	
	for(my $Row = 23; $Row <= 38; $Row++)   #NotesI_2
	{
	    my $name = &GetName( $Sheet->Cells($Row,"K")->{'Value'}, $diCount + 1);
	    if($name =~ /^(ScrubberOn)/)
		{
		   $name = "IntlkScrubberOn";
		}
	    printf OUT_FILE  "DEFINE_PLC_DI\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DI\\\\CH%d\", \t\"\\\\PLC\\\\DI\\\\SAFTETY\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	        "di".$name.",", $DeviceID, $BLOCK_SAFETY_IN, $index, $flag,
	        $ioOffset, $bitOffset, $Bits, $diCount, $name, "Safety DI";	
		$diCount++;
		$index++;
	    $bitOffset += $Bits;
	    if ($bitOffset >= $PLC_BITS) {
	    	$ioOffset += $bitOffset / $PLC_BITS;
	    	$bitOffset %= $PLC_BITS;
	    }
	}	
}

sub CreateSaftyOut
{
#     
#    printf OUT_FILE  "\r\n\/\/SafetyDI define\r\n";
    my $Sheet = $Book->Worksheets($Sheet_ProList);
	my $Bits = 1;

#  safty DO	
#    printf OUT_FILE  "\r\n\/\/SafetyDO define\r\n";
#	my $index = 272;
	my $index = 0;
	my $flag  = 0;
	my $ioOffset = 0;
	my $bitOffset = 0;

	for( my $Row = 4; $Row <= 19; $Row++)   #NotesI_1
	{
	    my $name = &GetName( $Sheet->Cells($Row,"N")->{'Value'}, $doCount + 1);
	    printf OUT_FILE  "DEFINE_PLC_DO\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DO\\\\CH%d\", \t\"\\\\PLC\\\\DO\\\\SAFTETY\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	        "do".$name.",", $DeviceID, $BLOCK_SAFETY_OUT, $index, $flag,
	        $ioOffset, $bitOffset, $Bits,$doCount+1, $name, "Safety DO";	
		$doCount++;
		$index++;
	    $bitOffset += $Bits;
	    if ($bitOffset >= $PLC_BITS) {
	    	$ioOffset += $bitOffset / $PLC_BITS;
	    	$bitOffset %= $PLC_BITS;
	    }
	}
	
	for(my $Row = 23; $Row <= 38; $Row++)   #NotesI_2
	{
	    my $name = &GetName( $Sheet->Cells($Row,"N")->{'Value'}, $doCount + 1);
	    printf OUT_FILE  "DEFINE_PLC_DO\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DO\\\\CH%d\", \t\"\\\\PLC\\\\DO\\\\SAFTETY\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	        "do".$name.",", $DeviceID, $BLOCK_SAFETY_OUT, $index, $flag,
	        $ioOffset, $bitOffset, $Bits,$doCount+1, $name, "Safety DO";	
		$doCount++;
		$index++;
	    $bitOffset += $Bits;
	    if ($bitOffset >= $PLC_BITS) {
	    	$ioOffset += $bitOffset / $PLC_BITS;
	    	$bitOffset %= $PLC_BITS;
	    }
	}		
}


sub CreateGasBoxIn
{
    my $ProfibusIn     = "B";
    my $ProfibusOut    = "F";
    my $Sheet = $Book->Worksheets($Sheet_ProList);
	
 #   printf OUT_FILE  "\r\n\/\/Gasbox input define\r\n";

	my $index = 0;
	my $flag  = 0;
	my $ioOffset = 0;
	my $bitOffset = 0;
    	
	#my $name = &GetName( $Sheet->Cells(150,$ProfibusIn)->{'Value'}, $aiCount + 1);
    #my $Bits = 32;
	#printf OUT_FILE  "DEFINE_PLC_IN_32BIT\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AI\\\\CH%d\", \t\"\\\\PLC\\\\AI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	#    "ai".$name.",", $DeviceID, $BLOCK_GASBOX_STATUS, $index, $flag,
	#    $ioOffset, $bitOffset, $Bits, $aiCount, $name, "IN_32BIT";	
	#$aiCount++;
    #$index++;
    #$bitOffset += $Bits;
    #if ($bitOffset >= $PLC_BITS) {
    #	$ioOffset += $bitOffset / $PLC_BITS;
    #	$bitOffset %= $PLC_BITS;
    #}
	
	my $Bits = 32; 		
	for( my $Row = 151; $Row <= 252; $Row = $Row+2)   #Gasbox Status
	{
	    if( $Row > 160)
	    {
		$flag =  $IO_FLAG_ASCII;
	    }
	    else
	    {
		$flag = 0;
	    }	
	
	    my $name = &GetName( $Sheet->Cells($Row,$ProfibusIn)->{'Value'}, $aiCount + 1);
	    printf OUT_FILE  "DEFINE_PLC_IN_32BIT\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AI\\\\CH%d\", \t\"\\\\PLC\\\\AI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	        "ai".$name.",", $DeviceID, $BLOCK_GASBOX_STATUS, $index, $flag,
	        $ioOffset, $bitOffset, $Bits, $aiCount, $name, "IN_32BIT";	
		$aiCount++;
		$index++;
	    $bitOffset += $Bits;
	    if ($bitOffset >= $PLC_BITS) {
	    	$ioOffset += $bitOffset / $PLC_BITS;
	    	$bitOffset %= $PLC_BITS;
	    }
	}
}



sub CreateGasBoxOut
{
    my $ProfibusIn     = "B";
    my $ProfibusOut    = "F";
    my $Sheet = $Book->Worksheets($Sheet_ProList);
	
  		
#   printf OUT_FILE  "\r\n\/\/Gasbox output define\r\n";

	my $index = 0;
	my $flag  = 0;
	my $ioOffset = 0;
	my $bitOffset = 0;	
		
	#my $Bits = 32;
	#my $name = &GetName( $Sheet->Cells(118,$ProfibusOut)->{'Value'}, $aoCount + 1);
	#printf OUT_FILE  "DEFINE_PLC_OUT_32BIT\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AO\\\\CH%d\", \t\"\\\\PLC\\\\AO\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	#    "ao".$name.",", $DeviceID, $BLOCK_GASBOX_CONTROL, $index, $flag,
	#    $ioOffset, $bitOffset, $Bits, $aoCount, $name, "OUT_32BIT";	
	#$aoCount++;
	#$index++;
    #$bitOffset += $Bits;
    #if ($bitOffset >= $PLC_BITS) {
    #	$ioOffset += $bitOffset / $PLC_BITS;
    #	$bitOffset %= $PLC_BITS;
    #}
	
    my $Bits = 32;
    my $Postfix = "OUT_32BIT";
    my $IOType = "DEFINE_PLC_OUT_32BIT";
        for( my $Row = 119; $Row <= 178; $Row = $Row+1)   #Gasbox control
	{
	    #printf OUT_FILE "|%s| %d\n",$Sheet->Cells($Row+1,$ProfibusOut)->{'Value'}, $Sheet->Cells($Row+1,$ProfibusOut)->{'Value'} eq "";
	    if($Sheet->Cells($Row+1,$ProfibusOut)->{'Value'} eq ""){
		$Bits = 32;
		$Postfix = "OUT_32BIT";
		$IOType = "DEFINE_PLC_OUT_32BIT";
	    }
	    else{
		#printf OUT_FILE "|%s|\n",$Sheet->Cells($Row+1,$ProfibusOut)->{'Value'};
		$Bits = 16;
		$Postfix = "OUT_16BIT";
		$IOType = "DEFINE_PLC_TEMPER_CONTROL";
	    }

	    my $name = &GetName( $Sheet->Cells($Row,$ProfibusOut)->{'Value'}, $aoCount + 1);
	    if( $Row > 132)
	    {
		$flag =  $IO_FLAG_ASCII;
	    }
	    else
	    {
		$flag = 0;
	    }

	    printf OUT_FILE  "%s\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AO\\\\CH%d\", \t\"\\\\PLC\\\\AO\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
	    $IOType, "ao".$name.",", $DeviceID, $BLOCK_GASBOX_CONTROL, $index, $flag,
	    $ioOffset, $bitOffset, $Bits, $aoCount+1, $name, $Postfix;	

	    if($Sheet->Cells($Row+1,$ProfibusOut)->{'Value'} eq ""){
		$Row = $Row + 1;
	    }

	    $aoCount++;
	    $index++;
	    $bitOffset += $Bits;
	    if ($bitOffset >= $PLC_BITS) {
		$ioOffset += $bitOffset / $PLC_BITS;
		$bitOffset %= $PLC_BITS;
	    }
	}	
}

sub CreateProfibusOut
{
    my $ProfibusIn     = "B";
    my $ProfibusOut    = "F";

    my $EndRowOfSafetyOut = 4;
    my $EndRowOfMotionOut = 64;       #PWO-61
    my $EndRowOfTemperOut = 118;      #PWO-115
    my $EndRowOfGasBoxOut = 178;      #PWO-175
	
    my $StartRowOfSafetyOut = 3;
    my $StartRowOfMotionOut = 3;
    my $StartRowOfTemperOut = 65;
    my $StartRowOfGasBoxOut = 117;	

    my $Sheet = $Book->Worksheets($Sheet_ProList);
    printf  DEBUG_FILE "Open\n";	

	my $ioOffsetTemper = 0;
	my $bitOffsetTemper = 0;
	my $ioOffsetMotion = 2;
	my $bitOffsetMotion = 0;
    my $Bits = 16;

	my $CHIndex = 0;
    foreach my $Row ($StartRow..$Sheet->{UsedRange}->{Rows}->{Count})
    {  	
	   my $flag  = 0;
       my $Name = &GetName( $Sheet->Cells($Row,$ProfibusOut)->{'Value'}, $aoCount + 1);
	   
       if($Row <= $EndRowOfSafetyOut)
	   {	   
	   }
	   elsif($Row <= $EndRowOfMotionOut)
	   {  
	      my $Index = $Row - $StartRowOfMotionOut;    #Safty/motion share
		  printf  DEBUG_FILE "%s \n",$Name;	 
 		  if($Index == 0)
		  {
              printf OUT_FILE  "\r\n\r\n";	 
#		  	  printf OUT_FILE  "\r\n\/\/Motion output define\r\n";
		  }
		  
		  printf OUT_FILE  "DEFINE_PLC_MOTION_CONTROL\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AO\\\\CH%d\", \t\"\\\\PLC\\\\AO\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
		      "ao".$Name.",", $DeviceID, $BLOCK_MOTION_CONTROL, $Index, $flag,
		      $ioOffsetMotion, $bitOffsetMotion, $Bits, $aoCount+1, $Name, "MOTION CONTROL"; 
		  $aoCount++;
		$bitOffsetMotion += $Bits;
		if ($bitOffsetMotion >= $PLC_BITS) {
			$ioOffsetMotion += $bitOffsetMotion / $PLC_BITS;
			$bitOffsetMotion %= $PLC_BITS;
		}
	   }
	   elsif($Row <= $EndRowOfTemperOut)
	   {
	      my $Index = $Row - $StartRowOfTemperOut;
		  printf  DEBUG_FILE "%s \n",$Name;	  
		  if($Index == 0)
		  {
		      printf OUT_FILE  "\r\n\r\n";	  
#		  	  printf OUT_FILE  "\r\n\/\/Motion output define\r\n";
		  }
		  
		  printf OUT_FILE  "DEFINE_PLC_TEMPER_CONTROL\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AO\\\\CH%d\", \t\"\\\\PLC\\\\AO\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
		  		"ao".$Name.",", $DeviceID, $BLOCK_TEMPER_CONTROL, $Index, $flag,
		  		$ioOffsetTemper, $bitOffsetTemper, $Bits,  $aoCount+1, $Name, "TEMPER CONTROL";    
		  $aoCount++;
		$bitOffsetTemper += $Bits;
		if ($bitOffsetTemper >= $PLC_BITS) {
			$ioOffsetTemper += $bitOffsetTemper / $PLC_BITS;
			$bitOffsetTemper %= $PLC_BITS;
		}
	   }
 
	   $CHIndex++;
   }	
}

sub CreateProfibusIn
{
    my $ProfibusIn     = "B";
    my $ProfibusOut    = "F";

    my $EndRowOfSafetyIn = 4;
    my $EndRowOfMotionIn = 56;    #PWI-53  for DWORD
    my $EndRowOfTemperIn = 150;   #PWI-145
    my $EndRowOfGasBoxIn = 273;   #PWI-269

    my $StartRowOfSafetyIn = 3;
    my $StartRowOfMotionIn = 3;
    my $StartRowOfTemperIn = 57;
    my $StartRowOfGasBoxIn = 149;	
	
    my $Sheet = $Book->Worksheets($Sheet_ProList);
    printf  DEBUG_FILE "Open\n";	
	my $CHIndex = 0;

	my $ioOffsetMotion = 2;
	my $bitOffsetMotion = 0;
	my $ioOffsetTemper = 0;
	my $bitOffsetTemper = 0;	
    my $Bits = 16;
	foreach my $Row ($StartRow..$Sheet->{UsedRange}->{Rows}->{Count})
    {  	
	   my $flag  = 0;
       my $Name = &GetName( $Sheet->Cells($Row,$ProfibusIn)->{'Value'}, $aiCount + 1);
	   
       if($Row <= $EndRowOfSafetyIn)
	   {	   
	   }
	   elsif($Row <= $EndRowOfMotionIn)
	   {  
	      my $Index = $Row - $StartRowOfMotionIn;  #Safty/motion share
		  printf  DEBUG_FILE "%s \n",$Name;	 		  
		  if($Index == 0)
		  {
		      printf OUT_FILE  "\r\n\r\n";	  
#		  	  printf OUT_FILE  "\r\n\/\/Motion input define\r\n";
		  }
		  
			printf OUT_FILE  "DEFINE_PLC_MOTION_STATUS\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AI\\\\CH%d\", \t\"\\\\PLC\\\\AI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
		    	"ai".$Name.",", $DeviceID, $BLOCK_MOTION_STATUS, $Index, $flag,
		    	$ioOffsetMotion, $bitOffsetMotion, $Bits, $aiCount+1, $Name, "MOTION STATUS"; 
			$aiCount++;
			$bitOffsetMotion += $Bits;
			if ($bitOffsetMotion >= $PLC_BITS) {
				$ioOffsetMotion += $bitOffsetMotion / $PLC_BITS;
				$bitOffsetMotion %= $PLC_BITS;
			}
	   }
	   elsif($Row <= $EndRowOfTemperIn)
	   {
	      my $Index = $Row - $StartRowOfTemperIn;
		  printf  DEBUG_FILE "%s \n",$Name;	 
		  if($Index == 0)
		  {
		      printf OUT_FILE  "\r\n\r\n";	  
#		  	  printf OUT_FILE  "\r\n\/\/Temper input define\r\n";
		  }
		  
		  $flag = $IO_FLAG_BIPOLAR;
		  #under table rule, aiMFCStatus flag set to 0
		  if ($Row == $EndRowOfTemperIn)
		  {
			$flag = 0;
		  }
		  
			printf OUT_FILE  "DEFINE_PLC_TEMPER_STATUS\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AI\\\\CH%d\", \t\"\\\\PLC\\\\AI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
		    	"ai".$Name.",", $DeviceID, $BLOCK_TEMPER_STATUS, $Index, $flag,
		    	$ioOffsetTemper, $bitOffsetTemper, $Bits, $aiCount+1, $Name, "TEMPER STATUS";
			$aiCount++;
			$bitOffsetTemper += $Bits;
			if ($bitOffsetTemper >= $PLC_BITS) {
				$ioOffsetTemper += $bitOffsetTemper / $PLC_BITS;
				$bitOffsetTemper %= $PLC_BITS;
			}
	   }
	   $CHIndex++;
    }  
}

sub CreateIO
{ 
   my $name  = $_[0];
   my $type  = $_[1];
   my $block = $_[2];
   my $index = $_[3];
	my $ioOffset = $_[5];
	my $bitOffset = $_[6];
   
   my $subtype = $_[4];
   
   my $flag  = 0;
   my $CHIndex   = $index + 1;

   my $Bits = 1;
   my $NameIsSpare = 0;
   if($name)
   {
		$name =~tr/(\(.*\))//d;
		$name =~tr/[ \r\n\t:\-]//d;
		$name =~tr/(\/)//d;

       if($name =~ /^(Spare|reserve|Reserve|spare)/)
       {
          $name = sprintf("CH%d", $CHIndex);
          $NameIsSpare = 1;
       }

       if($block == $BLOCK_DI)
       {          
          if($type eq "LocalDI")
          {
             $Bits = 1;
             printf OUT_FILE  "DEFINE_PLC_DI\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DI\\\\CH%d\", \t\"\\\\PLC\\\\DI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
                 "di".$name.",", $DeviceID, $block, $index, $flag,
                 $ioOffset, $bitOffset, $Bits,$CHIndex, $name, "DI";
          }
          elsif($type eq "SafetyDI")
          {
             $Bits = 1;
             printf OUT_FILE  "DEFINE_PLC_DI\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DI\\\\CH%d\", \t\"\\\\PLC\\\\DI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
                 "di".$name.",", $DeviceID, $block, $index, $flag,
                 $ioOffset, $bitOffset, $Bits,$CHIndex, $name, "Safety DI";
          }
          else
          {
             $name =~tr/[ \r\n\t]//d;
             $name =~s/(PM1|PMX|PMx|PM)//;
             if($NameIsSpare == 0 )
             {
               $type =~s/DI//;
               $name = $type.$name;
             }
             $Bits = 1; 
             printf OUT_FILE  "DEFINE_PLC_DI\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DI\\\\CH%d\", \t\"\\\\PLC\\\\DI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
                 "di".$name.",", $DeviceID, $block, $index, $flag,
                 $ioOffset, $bitOffset, $Bits, $CHIndex, $name, "DI";
          }
		  $diCount = $CHIndex;
       }
       elsif($block == $BLOCK_DO)
       {
          if($type eq "LocalDO")
          {
             $Bits = 1;
             printf OUT_FILE  "DEFINE_PLC_DO\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DO\\\\CH%d\", \t\"\\\\PLC\\\\DO\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
                 "do".$name.",", $DeviceID, $block, $index, $flag,
                 $ioOffset, $bitOffset, $Bits,$CHIndex, $name, "DO";
          }
          else
          {
             $name =~tr/[ \r\n\t]//d;
             $name =~s/(PM1|PMX|PMx|PM)//;
             if($NameIsSpare == 0 )
             {
                $type =~s/DO//;
                $name = $type.$name;
             } 
             $Bits = 1;
             printf OUT_FILE  "DEFINE_PLC_DO\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\DO\\\\CH%d\", \t\"\\\\PLC\\\\DO\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
                 "do".$name.",", $DeviceID, $block, $index, $flag,
                 $ioOffset, $bitOffset, $Bits,$CHIndex, $name, "DO";
          }
		  $doCount = $CHIndex;
       }
       elsif($block == $BLOCK_AI || $block == $BLOCK_SYSTEM_AI)
       {
			$name = GetName($_[0], $aiCount+1);
	   
          $Bits = 16; 
		  
		  
          if($type eq "SystemAI")
          {		
		  printf OUT_FILE  "DEFINE_PLC_IN_16BIT\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AI\\\\CH%d\", \t\"\\\\PLC\\\\AI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
              "ai".$name.",", $DeviceID, $block, $index, $flag,
              $ioOffset, $bitOffset, $Bits,$aiCount+1, $name, "System AI";		  
		  }
		  else
          {
		  printf OUT_FILE  "DEFINE_PLC_AI\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AI\\\\CH%d\", \t\"\\\\PLC\\\\AI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
              "ai".$name.",", $DeviceID, $block, $index, $flag,
              $ioOffset, $bitOffset, $Bits,$aiCount+1, $name, "AI";
   		  }
		  $aiCount++;                
       }  
       elsif($block == $BLOCK_AO || $block == $BLOCK_SYSTEM_AO)
       {
           $Bits = 16;  

          if($type eq "SystemAO")
          {		
           printf OUT_FILE  "DEFINE_PLC_OUT_16BIT\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AO\\\\CH%d\", \t\"\\\\PLC\\\\DI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
              "ao".$name.",", $DeviceID, $block, $index, $flag,
              $ioOffset, $bitOffset, $Bits, $CHIndex, $name, "System AO";		  
		  }
		  else
          {
           printf OUT_FILE  "DEFINE_PLC_AO\(%-50s  %d, %d, %3d, %d, %3d, %2d, %2d, \t\"\\\\PLC\\\\AO\\\\CH%d\", \t\"\\\\PLC\\\\DI\\\\%s\", \t\"CPU0\", \t\"PLC\", \t\"%s\"\)\n",
              "ao".$name.",", $DeviceID, $block, $index, $flag,
              $ioOffset, $bitOffset, $Bits, $CHIndex, $name, "AO";	
   		  }
		  

       }
   }
   return $Bits;
}


sub GetName
{
    my $name  = &TrimName($_[0]);
	my $index = $_[1];
	
    if($name && $name =~ /^(Spare|reserve|Reserve|spare)/)
    {
       $name = sprintf("CH%d", $index);
    }	
	
	return $name;	
}

sub TrimName
{
    my $name  = $_[0];
	if ($name) {
		$name =~tr/(\(.*\))//d;  
		$name =~tr/[ \r\n\t:\-]//d;
		$name =~tr/(\/)//d;
	}
	return $name;
}
