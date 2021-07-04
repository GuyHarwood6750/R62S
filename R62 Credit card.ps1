<#      Extract credit card transactions from spreadsheet to be processed as Pastel payment batch.
        Modify the $endR (endrow) as spreadsheet is added too.

#>
$inspreadsheet = 'C:\userdata\route 62\Bank Statements\CC\Credit Card transactions.xlsx'
$outfile2 = 'C:\userdata\route 62\Bank Statements\CC\CCTransactions_1.csv'
$custsheet = 'Transactions'                                #Transactions worksheet
$startR = 2                                             #Start row - do not change
$endR = 114                                              #End Row - change if necessary depending on number of transactions
$csvfile = 'SHEET1.csv'
$pathout = 'C:\userdata\route 62\Bank Statements\CC\'
$startCol = 1                                                                   #Start Col (don't change)
$endCol = 11                                                                     #End Col (don't change)
$filter = "P"                                                                   #Payments only
$outfile1 = 'C:\Userdata\route 62\Bank Statements\CC\CCTEMP.txt'              #Temp file
$outfileF = 'C:\Userdata\route 62\Bank Statements\CC\CCTransactionspastel.txt'             #File to be imported into Pastel             
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript { $_.P1 -eq $filter -and $_.P11 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

Get-ChildItem -Path $pathout -Name $csvfile
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($Outfile)
$xl.Sheets.Item('sheet1').Activate()
$range = $xl.Range("e:e").Entirecolumn
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
$data = Import-Csv -path $outfile2 -header type, GL, Expacc, ST, date, ref, desc, amt1, amt, vat     

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
    Switch ($aObj.Expacc) {
        ADV { $expacc = '3050000'; $aObj.descr }         
        Advertising { $expacc = '3050000'; $aObj.descr }         
        AIRTIME { $expacc = '4600000'; $aObj.descr }         
        BC { $expacc = '5201001'; $aObj.descr }             #Bank charges, private account      
        CLEANING { $expacc = '3250000'; $aObj.descr }         
        CWAGE { $expacc = '4401000'; $aObj.descr }         
        COMPA { $expacc = '6250010'; $aObj.descr }         
        COMPE { $expacc = '3300000'; $aObj.descr }         
        COURIER { $expacc = '3400000'; $aObj.descr }
        DONATION { $expacc = '3600000'; $aObj.descr }         
        DROT { $expacc = 'DROT'; $aObj.descr }         
        ELEC { $expacc = '3650000'; $aObj.descr }
        EQUIP { $expacc = '2999000'; $aObj.descr }
        FUEL { $expacc = '4150001'; $aObj.descr }
        INTP { $expacc = '5201001'; $aObj.descr }
        MVR { $expacc = '4150002'; $aObj.descr }
        OIL { $expacc = '4150001'; $aObj.descr }
        PACKAGING { $expacc = '2000010'; $aObj.descr }
        PC { $expacc = '8430000'; $aObj.descr }
        PCOMP { $expacc = 'PCOMP'; $aObj.descr }            #Supplier
        POST { $expacc = '3400000'; $aObj.descr }
        PVT { $expacc = '5201001'; $aObj.descr }
        PVTM { $expacc = '5201002'; $aObj.descr }
        PVTR { $expacc = '5201003'; $aObj.descr }
        NPUR { $expacc = '2000012'; $aObj.descr }
        PUR { $expacc = '2000010'; $aObj.descr }         
        RM { $expacc = '4350000'; $aObj.descr } 
        R62L { $expacc = '9994001'; $aObj.descr } 
        STAYC { $expacc = 'STAYC'; $aObj.descr }            #Supplier
        SYNT01 { $expacc = 'SYNT01'; $aObj.descr }          #Supplier
        STATIONERY { $expacc = '4200000'; $aObj.descr }
        TEL { $expacc = '4600000'; $aObj.descr }         
        UNIFORMS { $expacc = '4500000'; $aObj.descr }         
        WDEB { $expacc = 'WDEB'; $aObj.descr }         
        
        Default { $expacc = '9983000'; $aObj.descr }       
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
        GL      = $aObj.GL                      #GDC - general ledger, debtor, creditor
        contra  = $expacc                       #Expense account to be debited (DR)
        ref     = $aObj.ref
        comment = $aObj.desc
        amount  = $aObj.amt
        fil1    = $VATind
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '8420000'                     #Credit card contra account number
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