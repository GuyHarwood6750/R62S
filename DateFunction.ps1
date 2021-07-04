$mydate = [DateTime]::ParseExact('24/06/2021', 'dd/MM/yyyy', $null)
$year = $mydate.Year
$month = $mydate.Month
$lastday = [datetime]::DaysInMonth($year, $month)
$invoicedate = (Get-Date -Day $lastday -Month $month -Year $Year).ToString("dd/MM/yyyy")
Write-Output  $year $month $lastday $invoicedate