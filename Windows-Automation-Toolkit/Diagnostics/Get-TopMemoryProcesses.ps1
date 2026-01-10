# 1. Calculate and Display System Uptime
$OS = Get-CimInstance -ClassName Win32_OperatingSystem
$Uptime = (Get-Date) - $OS.LastBootUpTime

Write-Host "`n🕒 SYSTEM UPTIME REPORT" -ForegroundColor Cyan
Write-Host "--------------------------------" -ForegroundColor Gray
Write-Host "System has been UP for: $($Uptime.Days) Days, $($Uptime.Hours) Hours, $($Uptime.Minutes) Minutes" -ForegroundColor Green
Write-Host "Last Boot Time:         $($OS.LastBootUpTime)" -ForegroundColor Yellow

# 2. Top 10 Processes by Memory (Working Set)
Write-Host "`n💾 TOP 10 PROCESSES BY MEMORY USAGE" -ForegroundColor Cyan
Write-Host "---------------------------------------" -ForegroundColor Gray

Get-Process | 
    Sort-Object WS -Descending | 
    Select-Object -First 10 | 
    Select-Object Id, ProcessName, @{Name='Memory (MB)';Expression={[math]::Round($_.WS / 1MB, 2)}}, Path | 
    Format-Table -AutoSize