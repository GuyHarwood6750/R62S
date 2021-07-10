Function A1 {[CmdletBinding(PositionalBinding = $true)]
   param(
      [Parameter(Mandatory = $true)]
      [ValidateNotNull()]
      [String]$suppliername, [String]$month)
<# Move files            
 #>
   $dropboxBase = "C:\Users\Guy\Dropbox\R62\Accounts\"
   $year = "Suppliers 2022\"  
   $sourceAllFiles = $dropboxBase + $year + "_New invoices and credit notes\" + $month + '\'

Switch ($suppliername) {
   BLUE {
      $destFile = $dropboxBase + $year + "Blue Label\"
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceALLFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like 'Blue*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }  
   RENTOKIL {
      $destFile = $dropboxBase + $year + "Rentokil\"
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceALLFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like 'Rentokil*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }  
   PACKTOWN {
      $destFile = $dropboxBase + $year + "Pactown\"
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceALLFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like 'Packtown*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }  
   MIBCO {
      $destFile = $dropboxBase + $year + "Mibco\"
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceALLFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like 'Mibco*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }  
   "Cape Karoo" {
      $destFile = $dropboxBase + $year + "Cape Karoo International\"
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceALLFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like 'Cape Karoo*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }  
   "Star Card" {
      $destFile = $dropboxBase + "Star Cards\2022\"
      if (-Not (Test-Path -Path $destFile)) {
         Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
      }
      else {
         $bs = Get-ChildItem -Path $sourceALLFiles -file
         foreach ($bsf in $bs) {
            if ($bsf.name -like 'Star Card*.pdf') {
               $file = $bsf.Name
               Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
            }
            else {
            }
         }
      }                
   }
      JIREH {
         $destFile = $dropboxBase + $year + "Jireh\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like 'Jireh*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      TELKOM {
         $destFile = $dropboxBase + $year + "Telkom\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like 'Telkom*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      "Little Oaks" {
         $destFile = $dropboxBase + $year + "Little Oaks\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like 'Little Oaks*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      "Huge Connect" {
         $destFile = $dropboxBase + $year + "Huge Connectnet\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like 'Huge conn*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      "1-Grid" {
         $destFile = $dropboxBase + $year + "Gridhost\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like '1 Grid*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      Geiiansa {
         $destFile = $dropboxBase + $year + "Geiiansa\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like 'Geiiansa*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      SWD {
         $destFile = $dropboxBase + $year + "Swellendam Municipality\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like 'SWD*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      Sapphire {
         $destFile = $dropboxBase + $year + "Sapphire\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like 'Sapphire*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
      MOOV {
         $destFile = $dropboxBase + $year + "MOOV\"
         if (-Not (Test-Path -Path $destFile)) {
            Write-Error -Message "Folder does not exist '$destFile'. Error was: $_" -ErrorAction Stop
         }
         else {
            $bs = Get-ChildItem -Path $sourceALLFiles -file
            foreach ($bsf in $bs) {
               if ($bsf.name -like '*moov*.pdf') {
                  $file = $bsf.Name
                  Move-Item -Path $sourceALLFiles\$file -Destination $destFile -Force   
               }
               else {
               }
            }
         }                
      }    
}       
}
#A1 -suppliername 'RENTOKIL' -month '05July'
#A1 -suppliername 'packtown' -month '05July'
#A1 -suppliername 'mibco' -month '05July'
#A1 -suppliername 'cape karoo' -month '05July'
#A1 -suppliername 'blue' -month '05July'
#A1 -suppliername 'star card' -month '05July'
#A1 -suppliername 'jireh' -month '05July'
#A1 -suppliername 'telkom' -month '05July'
#A1 -suppliername 'little oaks' -month '05July'
#A1 -suppliername 'huge Connect' -month '05July'
#A1 -suppliername '1-Grid' -month '05July'
#A1 -suppliername 'Geiiansa' -month '05July'
#A1 -suppliername 'SWD' -month '05July'
#A1 -suppliername 'Sapphire' -month '05July'
#A1 -suppliername 'MOOV' -month '05July'