$c = VATCalc -amountincvat 80
Write-Host "Vat is $($c.vat) and amount excluding vat is $($c.vatexamt) and vat percent is $($c.vatpercent)"