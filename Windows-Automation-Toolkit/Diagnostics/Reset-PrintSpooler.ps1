Stop-Service -Name "Spooler" -Force
# Clear the spooler cache files
Remove-Item -Path "$env:SystemRoot\System32\spool\PRINTERS\*" -Force
Start-Service -Name "Spooler"
Write-Host "Print Spooler has been reset successfully." -ForegroundColor Green