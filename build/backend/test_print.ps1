$printer = 'Microsoft Print to PDF'
$path = Join-Path $env:TEMP 'teste_impressao.txt'
Set-Content -Path $path -Value 'Teste de impressao do backend`nData: 2026-08-10'
Write-Host "ARQUIVO: $path"
Write-Host "IMPRESSORA: $printer"
Get-Printer -Name $printer -ErrorAction Stop | Out-Null
Set-DefaultPrinter $printer
Start-Process notepad.exe -ArgumentList '/p', $path -WindowStyle Hidden -Wait
Start-Sleep -Seconds 5
Write-Host '--- FILA ---'
Get-PrintJob -PrinterName $printer -ErrorAction SilentlyContinue | Format-Table -AutoSize
Write-Host '--- TOTAL ---'
(Get-PrintJob -PrinterName $printer -ErrorAction SilentlyContinue | Measure-Object).Count
