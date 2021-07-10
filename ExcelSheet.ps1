function ExcelFormatDate {
  param ($file, $sheet, $column)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($file)
$xl.Sheets.Item($sheet).Activate()
$range = $xl.Range($column).Entirecolumn
$range.NumberFormat = 'dd/mm/yyyy'
$wb.save()
$xl.Workbooks.Close()
$xl.Quit()
}
ExcelFormatDate -file 'C:\Userdata\Route 62\_All Suppliers\test.csv' -sheet 'test' -column 'D:D'