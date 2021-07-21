Function A1A {[CmdletBinding(PositionalBinding = $true)]
   param(
      [Parameter(Mandatory = $true)]
      [ValidateNotNull()]
      [String]$reportname)
<# Move files            
 #>
   $BASEEasypay = "C:\Userdata\Route 62\EasyPay\"
   $BASEEOS = "C:\Userdata\Route 62\EOS Not Processed\Reports\"
   $year = "2022\"  
   $sourceEasypayFiles = $BASEEasypay
   $sourceEOSFiles = $BASEEOS

Switch ($reportname) {
   EASYPAY {
      $destFile = $BASEEasypay + 'Completed\' + $year
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceEasypayFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like '*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceEasypayFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }  
   EOS {
      $destFile = $BASEEOS + 'Processed\' + $year
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceEOSFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like '*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceEOSFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }   
}       
}
A1A -reportname 'EasyPay'
#A1A -reportname 'EOS'
