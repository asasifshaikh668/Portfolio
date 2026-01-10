function Repair-CriticalServices {
    param (
        [string[]]$ServiceNames = @("Spooler", "wuauserv", "Dnscache", "CryptSvc")
    )

    foreach ($Name in $ServiceNames) {
        # Check if service exists
        $Service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $Service) {
            Write-Warning "Service '$Name' not found."
            continue
        }

        # Check if NOT Running (Stopped, Paused, or Stuck)
        if ($Service.Status -ne 'Running') {
            # Alert the user
            Write-Host "⚠️  ALERT: Service '$Name' is currently [$($Service.Status)]" -ForegroundColor Yellow
            
            # Prompt for action
            $UserChoice = Read-Host "   Do you want to start this service? (Y/N)"

            if ($UserChoice -eq 'Y') {
                Write-Host "   Starting '$Name'..." -NoNewline
                try {
                    Start-Service -Name $Name -ErrorAction Stop
                    Write-Host " [SUCCESS] ✅" -ForegroundColor Green
                }
                catch {
                    Write-Host " [FAILED] ❌" -ForegroundColor Red
                    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "   Skipping '$Name'." -ForegroundColor Gray
            }
        }
        else {
            Write-Host "✅ Service '$Name' is running normally." -ForegroundColor Green
        }
    }
}

# Run the function
Repair-CriticalServices