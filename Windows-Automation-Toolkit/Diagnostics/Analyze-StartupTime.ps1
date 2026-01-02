function Analyze-StartupTime {
    # Define Event IDs for Boot (6005) and Shutdown (6006, 6008, 1074)
    $EventIDs = @(6005, 6006, 6008, 1074)
    
    Write-Host "Retrieving recent Boot and Shutdown events..." -ForegroundColor Cyan
    
    # Query the System Event Log for these specific IDs
    Get-WinEvent -FilterHashtable @{
        LogName = 'System'
        Id = $EventIDs
    } -MaxEvents 20 | Select-Object TimeCreated, Id, @{Name='Type';Expression={
        if ($_.Id -eq 6005) { "Boot/Startup" }
        elseif ($_.Id -eq 6008) { "Unexpected Shutdown" }
        else { "Normal Shutdown" }
    }}, Message | Format-Table -AutoSize
}

Analyze-StartupTime