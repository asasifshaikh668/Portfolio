# Get-AllUserLoginHistory.ps1

param (
    [int]$MaxEvents = 50
)

# 4624 = Success, 4625 = Failure
$EventIDs = @(4624, 4625) 

Write-Host "Retrieving last $MaxEvents login events for ALL users..." -ForegroundColor Cyan

Get-WinEvent -FilterHashtable @{
    LogName = 'Security'
    ID      = $EventIDs
} -MaxEvents $MaxEvents -ErrorAction SilentlyContinue | Select-Object TimeCreated, Id, @{Name='Status';Expression={
    if ($_.Id -eq 4624) { "Success" } else { "FAILED" }
}}, @{Name='User';Expression={
    # Validates if the user property exists (Index 5 is usually TargetUserName)
    if ($_.Properties[5].Value) { $_.Properties[5].Value } else { "System/Unknown" }
}} | Format-Table -AutoSize