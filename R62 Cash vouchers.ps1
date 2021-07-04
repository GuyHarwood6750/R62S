<#      Extract cash vouchers from spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
                This can only be done by eyeball as spreadsheet has historical data.
#>
$inspreadsheet = 'C:\userdata\route 62\_All Suppliers\Suppliers JUNE 2021.xlsm'
$outfile2 = 'C:\userdata\route 62\_All Suppliers\CSH JUNE 2021_1.csv'
$custsheet = 'JUNE 2021'                                #Month worksheet
$startR = 5                                             #Start row - do not change
$endR = 204                                             #End Row - change if necessary depending on number of purchases
$csvfile = 'SHEET1.csv'
$pathout = 'C:\userdata\route 62\_All Suppliers\'
$startCol = 1                                                                   #Start Col (don't change)
$endCol = 11                                                                     #End Col (don't change)
$filter = "CSH"
$outfile1 = 'C:\Userdata\route 62\_all suppliers\cashsupplier.txt'              #Temp file
$outfileF = 'C:\Userdata\route 62\_all suppliers\cashpurpastel.txt'             #File to be imported into Pastel             
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript { $_.P2 -eq $filter -and $_.P11 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

Get-ChildItem -Path $pathout -Name $csvfile
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($Outfile)
$xl.Sheets.Item('sheet1').Activate()
$range = $xl.Range("d:d").Entirecolumn
$range.NumberFormat = 'dd/mm/yyyy'

$wb.save()
$xl.Workbooks.Close()
$xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile

#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $outfile2 -header Expacc, type, supplier, date, ref, ref2, descr, amt, vat    

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
    Switch ($aObj.Expacc) {
        ADV { $expacc = '3050000'; $description = $aObj.descr }         
        Advertising { $expacc = '3050000'; $description = $aObj.descr }         
        AIRTIME { $expacc = '4600000'; $description = $aObj.descr }         
        CLEANING { $expacc = '3250000'; $description = $aObj.descr }         
        CWAGE { $expacc = '4401000'; $description = $aObj.descr }         
        COMPA { $expacc = '6250010'; $description = $aObj.descr }         
        COMPE { $expacc = '3300000'; $description = $aObj.descr }         
        COURIER { $expacc = '3400000'; $description = $aObj.descr }
        DONATION { $expacc = '3600000'; $description = $aObj.descr }         
        ELEC { $expacc = '3650000'; $description = $aObj.descr }
        EQUIP { $expacc = '2999000'; $description = $aObj.descr }
        FUEL { $expacc = '4150001'; $description = $aObj.descr }
        GENERATOR { $expacc = '3650000'; $description = $aObj.descr }
        MVR { $expacc = '4150002'; $description = $aObj.descr }
        OIL { $expacc = '4150001'; $description = $aObj.descr }
        PACKAGING { $expacc = '2000010'; $description = $aObj.descr }
        POST { $expacc = '3400000'; $description = $aObj.descr }
        NPUR { $expacc = '2000012'; $description = $aObj.descr }
        PUR { $expacc = '2000010'; $description = $aObj.descr }         
        PVT { $expacc = '5201001'; $description = $aObj.descr }         
        RM { $expacc = '4350000'; $description = $aObj.descr } 
        STATIONERY { $expacc = '4200000'; $description = $aObj.descr }
        TEL { $expacc = '4600000'; $description = $aObj.descr }         
        UNIFORMS { $expacc = '4500000'; $description = $aObj.descr }         
        
        Default { $expacc = '9983000'; $description = $aObj.descr }       
    }

    Switch ($aObj.vat) {
        Y { $VATind = '15' }
        N { $VATind = '0' }
        Default {$VATind = '15'}
    }
    #Format Pastel batch   
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = 'G'
        contra  = $expacc                       #Expense account to be debited (DR)
        ref     = $aObj.ref
        comment = $description
        amount  = $aObj.amt
        fil1    = $VATind
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '8430000'                     #Cash voucher contra account number
        fil6    = '1'
        fil7    = '1'
        fil8    = '0'
        fil9    = '0'
        fil10   = '0'
        amt2    = $aObj.amt
    }
      
        $objlist = New-Object -TypeName psobject -Property $props1
        $objlist | Select-Object * | Export-Csv -path $outfile1 -NoTypeInformation -Append
    }  
    #Remove header information so file can be imported into Pastel Accounting.
    Get-Content -Path $outfile1 | Select-Object -skip 1 | Set-Content -path $outfilef
    Remove-Item -Path $outfile1