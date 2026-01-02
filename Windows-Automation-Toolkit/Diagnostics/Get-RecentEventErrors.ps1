function Get-EventLogErrors {
    $LogNames = @("System", "Application")
    
    foreach ($Log in $LogNames) {
        Write-Host "--- Last 50 Errors/Critical in $Log ---" -ForegroundColor Cyan
        # Using Get-WinEvent to capture Critical (1) and Error (2) levels
        Get-WinEvent -FilterHashtable @{
            LogName = $Log
            Level = 1, 2
        } -MaxEvents 50 -ErrorAction SilentlyContinue | 
        Select-Object TimeCreated, Id, LevelDisplayName, Message | 
        Format-Table -AutoSize
    }
}

Get-EventLogErrors