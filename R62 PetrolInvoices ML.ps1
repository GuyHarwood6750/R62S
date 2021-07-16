<#  Get list of petrol invoices from Petrol books spreadsheet
    +++++++++++++ MULTILINE invoices ++++++++++++
    Output to text file to be imported as a Pastel Invoice batch.

#>
$csvclient = 'C:\userdata\route 62\petrol books\DEKB.csv'      #Input from Client spreadsheet
$outfile = 'C:\userdata\route 62\petrol books\petrolinv.txt'        #Temp file
$outfile2 = 'C:\userdata\route 62\petrol books\DEKBpastel.txt'     #File to be imported into Pastel

#Remove last file imported to Pastel

$checkfile = Test-Path $outfile2
if ($checkfile) { Remove-Item $outfile2 }                   

#Import latest csv from Client spreadsheet

$data = Import-Csv -path $csvclient -header acc, date, invnum, ordernum, reg, lt, fuel, amt, slipno

$previnvnum = 0

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

if ($aObj.invnum -ne $previnvnum) {
    $previnvnum = $aObj.invnum
    #Format Pastel batch
    $header = [ordered] @{
        col1    = 'Header'
        col2    = ''
        col3    = ''
        col4    = 'Y'
        col5   = $aObj.acc
        col6   = $pastelper
        col7   = $aObj.date
        col8 = $aObj.invnum
        col9    = "N"
        col10    = '0'
        col11    = ''
        col12    = ''
        col13    = ''
        col14    = ''
        col15   = ''
        col16   = ''
        col17   = ''
        col18   = ''
        col19   = ''
        col20   = '0'
        col21   = $aObj.date
        col22   = ''
        col23   = ''
        col24   = ''
        col25   = '1'
        col26   = ''
        col27   = ''
        col28   = ''
        col29   = ''
    }
    $objlist = New-Object -TypeName psobject -Property $header 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
}    
    $details1 = [ordered] @{
        col1   = 'Detail'
        col2   = '0'
        col3   = '1'
        col4   = $aObj.amt
        col5   = $aObj.amt
        col6   = ''
        col7   = '0'
        col8   = '0'
        col9   = '0'
        col10   = '8430000'
        col11   = 'Product' + ' : ' + $aObj.fuel
        col12   = '6'
        col13   = ''
        col14   = ''
        col15   = ''
        col16   = ''
        col17   = ''
        col18   = ''
        col19   = ''
        col20   = ''
        col21   = ''
        col22   = ''
        col23   = ''
        col24   = ''
        col25   = ''
        col26   = ''
        col27   = ''
        col28   = ''
        col29   = ''
    }
    $details2 = [ordered] @{
        col1   = 'Detail'
        col2   = '0'
        col3   = '1'
        col4   = '0'
        col5   = '0'
        col6   = ''
        col7   = '0'
        col8   = '0'
        col9   = '0'
        col10   = "'"
        col11   = 'Lt' + ' : ' + $aObj.lt
        col12   = 7
        col13   = ''
        col14   = ''
        col15   = ''
        col16   = ''
        col17   = ''
        col18   = ''
        col19   = ''
        col20   = ''
        col21   = ''
        col22   = ''
        col23   = ''
        col24   = ''
        col25   = ''
        col26   = ''
        col27   = ''
        col28   = ''
        col29   = ''
    }
    $details3 = [ordered] @{
        col1   = 'Detail'
        col2   = '0'
        col3   = '1'
        col4   = '0'
        col5   = '0'
        col6   = ''
        col7   = '0'
        col8   = '0'
        col9   = '0'
        col10   = "'"
        col11   = 'Order' + ' : ' + $aObj.ordernum
        col12   = 7
        col13   = ''
        col14   = ''
        col15   = ''
        col16   = ''
        col17   = ''
        col18   = ''
        col19   = ''
        col20   = ''
        col21   = ''
        col22   = ''
        col23   = ''
        col24   = ''
        col25   = ''
        col26   = ''
        col27   = ''
        col28   = ''
        col29   = ''
    }
    $details4 = [ordered] @{
        col1   = 'Detail'
        col2   = '0'
        col3   = '1'
        col4   = '0'
        col5   = '0'
        col6   = ''
        col7   = '0'
        col8   = '0'
        col9   = '0'
        col10   = "'"
        col11   = 'Reg' + ' : ' + $aObj.Reg
        col12   = 7
        col13   = ''
        col14   = ''
        col15   = ''
        col16   = ''
        col17   = ''
        col18   = ''
        col19   = ''
        col20   = ''
        col21   = ''
        col22   = ''
        col23   = ''
        col24   = ''
        col25   = ''
        col26   = ''
        col27   = ''
        col28   = ''
        col29   = ''
    }
    $details5 = [ordered] @{
        col1   = 'Detail'
        col2   = '0'
        col3   = '1'
        col4   = '0'
        col5   = '0'
        col6   = ''
        col7   = '0'
        col8   = '0'
        col9   = '0'
        col10   = "'"
        col11   = 'Slip no' + ' : ' + $aObj.slipno
        col12   = 7
        col13   = ''
        col14   = ''
        col15   = ''
        col16   = ''
        col17   = ''
        col18   = ''
        col19   = ''
        col20   = ''
        col21   = ''
        col22   = ''
        col23   = ''
        col24   = ''
        col25   = ''
        col26   = ''
        col27   = ''
        col28   = ''
        col29   = ''
    }                   
    $objlist = New-Object -TypeName psobject -Property $details1 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
    $objlist = New-Object -TypeName psobject -Property $details2 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
    $objlist = New-Object -TypeName psobject -Property $details3 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
    $objlist = New-Object -TypeName psobject -Property $details4 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
    $objlist = New-Object -TypeName psobject -Property $details5 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
}    
#Remove header information so file can be imported into Pastel Accounting.

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile