Clear-Host
#$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
# Functions need to be defined b/f the rest of the code.
#
# 
#
# when a block is copied from Excel columns are seperated 
# with 09 (tab),  rows are seperated with  0x0D 0x0A (CR,LF)
#
# ------------
# 
# Get-Clipboard -Raw     will be used to keep each line from
# being it's onwn objest.  Later .split will be used.
#
# ------------
#
#
#
# Also the last column of the last row finishes with an additional
# 0D 0A  (CR,LF) combination.  0D is decimal 13   0A is decimal 10.
#
#   For following four cells of Excel data copied to the clipboard.
#    10      71
#    20      88
#  copied to clipboard and set to $t
#$t = Get-Clipboard -Format Text   # grab info on clipboard  
#$t  | Out-String |Format-Hex
$out=@'

           00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F

00000000   31 30 09 37 31 0D 0A 32 30 09 38 38 0D 0A 0D 0A  10.71..20.88....
'@
# Note the doble set of   0A 0D    at the end. 

# if a config file does not exist make one
if ( !(Test-Path -path C:\tmp\Make-uStation.cfg)  ) {

$uStation_cfg = @"
{
    "ColumnSpaArray":  [

                       ],
    "col_spa_code":  [
                         1
                     ],
    "line_spacing":  0.25,
    "justification_pat_code":  [
                                   "CC"
                               ],
    "dx":  [

           ],
    "ColumnJustArray":  [

                        ]
}
"@
   Set-Content -path C:\tmp\Make_uStation.cfg $uStation_cfg
}


Function Get-ClipboardMatrix  {
# There are two sets of CRLF at the end so we drop one of te array rows
# at the end
  Param ( $t )

    $array= @()  #empty list 
    $cnt=0  
    $t.split("`n`r")  |  Foreach -Process  {
      $row = $_
      #write-host $_ 
      $col = 0  
      $array += ,@()  #append an empty list resulting in list of lists
      
      $row.split("`t") | Foreach -Process {
          $cellval = $_ -replace "`r",""  #drop any esaped returns
          # Write-Host $cellval
          $array[$cnt] += $cellval  #add the next cell
          $col++
        }
      
      $cnt++
    }
    $last = $array.Count - 2  #drop the last one  
    $array = $array[0 .. $last]
    return $array
}  

Function Get-ColumnSpaArray {   
# usage:
# usage Get-ColumnSpaArray  -code $<>  -n $<>l 
# return an array with the column spacings 
# $col_spa_code is an array to be repeated as needed unless
# the last value is ".."  in that case the previous spacing is 
# to be repeated


  Param ($code=@(1), [int]$n)
  $n--  #one less spacings than columns in the matrix
  if ( $code[-1] -eq "..") {        
    [Int]$last = $code.count - 1   #what is the last index for $col_spa goint to be?
    # fill in the front part starting with a zero to be replaced later
    # by the negative sum of the other column spacings
    $col_spa = @( ,0 + $code[ 0 .. ($last - 1)] )  
    
      for( $j=$last-1;  $j -le $n-2  ; $j++ ) {        # pad the remaining. 
        #write-host "j = $j   n=$n  last=$last"
        $col_spa += $code[-2]
      }
    }
  else {
     $col_spa = @( ,0 )
     for( $j=0;  $j -le $n-1 ; $j++ ) {
       $col_spa += $code[ $j % $code.count ] 
       }
     }
  #write-host "col_spa=" $col_spa
  $col_spa[1 .. $n] | %  {$nsum -= $_ }
  $col_spa[0] = $nsum
  return $col_spa
}

Function Get-ColJustArray {   
# usage:
# Get-ColJustArray  -code $<>   -n $<> 
# return an array with the column spacings 
# $col_spa_code is an array to be repeated as needed unless
# the last value is ".."  in that case the previous spacing is 
# to be repeatedt
    Param ($code=@(1), [int]$n)
    $j_array = @(0 .. $($n-1) )  #empty array of correct length
    #error text 
    $str = @"

First code not valid
Try one of:
 
LT   CT   RT
LC   CC   RC
LB   CB   RB

or 

..  to repeat the previous code

Current Code array:
$code

"@
    #validate the codes    
    if( [bool]($code.ToUpper() -notmatch "[LCR][TCB]|\.\." ) -or
        $code[0].Toupper() -notmatch  "[LCR][TCB]" ) {
        Clear-Host
        Write-Host 
        Write-Host $str
        Read-Host -Prompt  "Script will stop" 
        Stop
        }
    # two ways of filling  ..
    if ( $code[-1] -eq "..") {     #fill the front then pad with the last known                       
        
        #all member up to the ".." become array "low"
        $low = @($code).Where( { $_ -eq ".." }, 'Until' ) 

        #write-host "Low:  $low" 
        # make a "high" array with a copy of the last member from "low" 
        # j_array already is the correct size. 
        ForEach ($i in $j_array[ ($low.count) .. ($j_array.count-1)] ) { 
            $high += ,$low[-1]  # note the array construction comma 
        }
        $j_array = $low + $high
    }
    else {
        $counter = 0
        ForEach ($i in $j_array) { 
            $j_array[$counter]= $code[ $counter % $code.count ]
            $counter++
        } 
    }
    #ConvertTo-Json -InputObject $j_array
    return $j_array
}

Function Get-dx ($inarray) {
    $dx = [ordered]@{}
    #Convertto-JSON -InputObject $inarray
    $last = $inarray.count - 1

    ForEach ( $i in 0 .. $last )  {
        $sum = 0.0
        $j =  ($i + 1)  %  $inarray.count 
        ForEach ( $x in  $j .. ($j + $inarray.count) ) {
            $x_current = $x % $inarray.count  # b/w 0 and $last
            #Write-host -NoNewline "$i, $j   "
            $sum += $inarray[$x_current]
            $dx."$i, $x_current" = [Math]::Round( $sum, 8 )
        }
        
    }
    return $dx
}

Function Get-uStationCommands  {
    Param ( $ValueMatrix, $config )
    $dx = Get-dx -inarray $config.ColumnSpaArray
    $justification_old = ""
    $row = 0
    # note ForEach-Object not exactly the same at ForEach
    $ValueMatrix | ForEach-Object  -Begin {
        
        [string]$outstr = ""
        #$outstr += "reset`r`n" 
        #$outstr += "place text`r`n" 
        } -Process {        
        $temp = $_
        $col = 0
        # note ForEach-Object not exactly the same at ForEach
        $temp | ForEach-Object  -BEGIN {
            } -Process {
            $justification_new = $config.ColumnJustArray[$col]
            #write-host " here "
            if ($_ -ne "" ) {  

                if ( $justification_old -ne $justification_new ) {
                    $outstr += "reset`r`n"
                    $outstr += "active txj $justification_new`r`n"
                    $outstr += "place text`r`n"
                    $justification_old = $justification_new
                }
 
                $outstr += "``$_`r`n"
                if ($row -eq 0 -and  $col -eq 0 ) {
                    $outstr += "dx=0,0`r`n" 
                    }

                else { 
                       # do not leave a spaces in the "=x,y" section otherwise
                       # microstation will not recognize the negative numbers correctly 
                       #write-host -Separator "" -NoNewline  "dx= " ($dx."$last_col, $col") " , " 
                       $outstr +=   "dx={0:n8}," -f   $dx."$last_col, $col"
                       $outstr +=   "{0:n8}`r`n"     -f   (($row - $last_row) * -$config.line_spacing)
                       
                }
                $last_col = $col
                $last_row = $row
            }    
            $col++
            }  -END { 
        }
        write-host   #blank line at end of columns
        $row++
        } -END { $outstr +=  "reset`r`n`r`n"
        }
        #write-host $outstr
        #$outstr |Get-Member
        return $outstr
    }

Function Loop-Matrix ( $inMatrix ) {
    if ( $true ) {
        #$inMatrix.GetType()
        $inMatrix | Foreach {
           $temp = $_
           $temp | Foreach {
              write-host -NoNewline $_ " " 
           }
           write-host
           }

        ConvertTo-Json $array
    }
}

Function Get-LineSpacing {
    param ($config )  
    [decimal]$lineSpa = Read-Host -prompt "`n`nEnter a line spacing in master units"
    write-host $lineSpa
    return $lineSpa

}



$t = Get-Clipboard -Format Text   # grab info on clipboard  
#$t  | out-string |Format-Hex
$matrix = Get-ClipboardMatrix $t 

$columns = $matrix[0].count



$config = @{       line_spacing = 0.25;             
         justification_pat_code = @( "CC"  ); 
                   col_spa_code = @( 1   ) ; 
                          
                           }




Function Get-Menu {

$config.ColumnSpaArray  = &Get-ColumnSpaArray -code $config.col_spa_code           -n $matrix[0].Count
$config.ColumnJustArray = &Get-ColJustArray   -code $config.justification_pat_code -n $matrix[0].Count
$config.dx = &Get-dx -inarray $config.ColumnSpaArray 

#$metta.ColumnSpaArray  = &Get-ColumnSpaArray -code $config.col_spa_code           -n $matrix[0].Count
#$metta.ColumnSpaArray = &Get-ColumnSpaArray -code $config.col_spa_code           -n $matrix[0].Count

return $menu = @"

1). Change Justification _Code_:    $($config.justification_pat_code) 
                  Justification:    $(&Get-ColJustArray -code $config.justification_pat_code -n $matrix[0].Count)


2). Change Column Spacing _Code_:   $($config.Col_spa_code)
                      Column Spa:   $(&Get-ColumnSpaArray -code $config.col_spa_code -n $matrix[0].Count)


3). Change Line Spacing:            $($config.line_spacing)

4). Save current config

5). Read config from c:\tmp\Make-uStation.cfg 

6). Create uStation commands

7 or E).  EXIT


"@
}



#ConvertTo-Json $matrix
#ConvertTo-JSON $config

while ($true) {
Clear-host
write-host "Number of columns    =  " $matrix[0].Count
Write-host "Number of Rows       =  " $matrix.Count
write-host "First row            =  " $matrix[0]
write-host "Last row             =  " $matrix[-1]

#$m = Get-Menu
$action = Read-Host -Prompt "$(&Get-Menu)`n`n"
Write-host $action


if ($action -eq 1) {
    Write-host
    Write-Host 'Enter new values like: "CC  RC .. '
    Write-host '<space seperated without quotes>'
    # mostly working do similar for action 2
    $justCode = Read-Host 'Input:  '
 #   $justCode.GetType()
    $justCode= $justCode.trim().ToUpper()
    $justCode= $justCode -replace ","," " -replace "\s+"," "
    $justCode = $justCode.Split(" ")
    $config.justification_pat_code = $justCode
    #Write-host $config.justification_pat_code
    #$config.justification_pat_code | convertTo-JSON -Compress
#    Pause
    
}

if ($action -eq 2) {
    Write-host
    Write-Host 'Enter new values like: "1.2   3.4   .. '
    Write-host '<space seperated without quotes>'
    # mostly working do similar for action 2
    $spaCode = Read-Host 'Input:  '
 #   $justCode.GetType()
    $spaCode = $spaCode.trim()
    $spaCode = $spaCode -replace ","," " -replace "\s+"," "
    $spaCode = $spaCode.Split(" ")
    @( $spaCode )  
    $config.col_spa_code = $spaCode
#    Pause
}


if ($action -eq 3) { $config.line_spacing = &Get-LineSpacing}

if ($action -eq 5) { $config = Get-Content C:\tmp\Make-uStation.cfg |ConvertFrom-Json}

if ($action -eq 4) {
    $config.ColumnSpaArray = @()
    $config.dx = @()
    $conifg.ColumnJustArray = @()

    $config | ConvertTo-Json | Out-File  C:\tmp\Make-uStation.cfg  -Encoding ascii

}

if ($action -eq 6) { 
    $Out = &Get-uStationCommands -ValueMatrix $matrix -config $config
    Write-host "---`n$Out`n---"
    $file = "c:\tmp\ustation.txt"
    $outfile = New-Item -path ($file) -ItemType file -Force 
    $Out | out-file -Encoding ascii -Append    $outfile
    Write-host "file written" 
    #Pause
    stop-process -Id $PID 
    Exit
}
if ($action -eq 7 -or $action.ToUpper() -eq "E") { 
    stop-process -Id $PID
    Exit 
    }



}

