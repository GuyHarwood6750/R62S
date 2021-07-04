<#      Extract from Municipality spreadsheet the range for new invoices to be generated.
        Modify the $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\route 62\_All Customers\Swellendam_Mun.xlsx'
$csvfile = 'sheet1.csv'
$pathout = 'C:\userdata\route 62\_All Customers\'
$custsheet = 'sheet1'                                   #Municipality worksheet
$startR = 1                                             #Start row (don't change)
$endR = 21                                              #End Row
$startCol = 1                                           #Start Col (don't change)
$endCol = 8                                             #End Col (don't change)

$Outfile = $pathout + $csvfile
$outfile2 = 'C:\userdata\route 62\_all Customers\municipality.csv' 

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript {$_.P8 -ne 'done' } | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('sheet1').Activate()
  
        $range = $xl.Range("b:b").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 2 | Set-Content -path $outfile2
Remove-Item -Path $outfile

$outfile3 = 'C:\userdata\route 62\_All Customers\temp.csv'        #Temp file
$outfile4 = 'C:\userdata\route 62\_All Customers\municipalityPastel.txt'     #File to be imported into Pastel

#Remove last file imported to Pastel

$checkfile = Test-Path $outfile4
if ($checkfile) { Remove-Item $outfile4 }                   

#Import latest csv from Client spreadsheet

$data = Import-Csv -path $outfile2 -header acc, date, invnum, ordernum, desc, amt, slipno

$previnvnum = 0

foreach ($aObj in $data) {
    #Return Pastel accounting period and Last day of month based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    $invoicedate = LastDayofMonth -CurrentDay $aObj.date
    
if ($aObj.invnum -ne $previnvnum) {
    $previnvnum = $aObj.invnum

    #Format Pastel batch
    $header = [ordered] @{
        Col1    = 'Header'
        Col2    = ''
        Col3   = ''
        Col4    = 'Y'
        Col5   = $aObj.acc
        Col6   = $pastelper
        Col7   = $invoicedate
        Col8 = $aObj.invnum
        Col9    = "Y"               #inclusive vat indicator
        Col10    = '0'
        Col11    = ''
        Col12    = ''
        Col13    = ''
        Col14    = ''
        Col15   = ''
        Col16   = ''
        Col17   = ''
        Col18   = ''
        Col19   = ''
        Col20   = '0'
        Col21   = $invoicedate
        Col22   = ''
        Col23   = ''
        Col24   = ''
        Col25   = '1'
        Col26   = ''            
        Col27   = ''            
        Col28   = ''            
        Col29   = ''            
    }
    $objlist = New-Object -TypeName psobject -Property $header 
    $objlist | Select-Object * | Export-Csv -path $outfile3 -NoTypeInformation -Append
}
    [decimal]$amount = $aObj.amt  
    [decimal]$vat = $amount * 15 / 115
    [decimal]$amtexvat = $aObj.amt - $vat
    $vatexamt = [math]::Round($amtexvat, 2)
    $vatpercent = 15

        $details1 = [ordered] @{
        Col1   = 'Detail'
        Col2   = '0'
        Col3   = '1'
        Col4   = $vatexamt
        Col5   = $aObj.amt
        Col6   = ''
        Col7   = $vatpercent
        Col8   = '0'
        Col9   = '0'
        Col10   = '1010100'
        Col11   = $aObj.desc + ' : ' + $aObj.slipno + ' : ' + $aObj.date
        Col12   = '6'
        Col13   = ''
        Col14   = ''
        Col15   = ''
        Col16   = ''
        Col17   = ''
        Col18   = ''
        Col19   = ''
        Col20   = ''
        Col21   = ''
        Col22   = ''
        Col23   = ''
        Col24   = ''
        Col25   = ''
        Col26   = ''
        Col27   = ''
        Col28   = ''
        Col29   = ''
    }   
    $objlist = New-Object -TypeName psobject -Property $details1 
    $objlist | Select-Object * | Export-Csv -path $outfile3 -NoTypeInformation -Append
}  
#Remove header information so file can be imported into Pastel Accounting.

Get-Content -Path $outfile3 | Select-Object -skip 1 | Set-Content -path $outfile4
Remove-Item -Path $outfile3
Remove-Item -Path $outfile2