Get-CPUTemperature {
    $temperatureData = Get-WmiObject -Class Win32_PerfFormattedData_Counters_ThermalZoneInformation -Namespace root/CIMV2 -ErrorAction SilentlyContinue

    Write-Host "`n--- 💡 CPU Temperature Check ---`n" -ForegroundColor Yellow

    if (-not $temperatureData) {
        Write-Host "❗ WARNING: Temperature data is not available on this system." -ForegroundColor Red
        Write-Host "This usually means the BIOS or WMI does not expose the sensor data." -ForegroundColor White
        return
    }
    
    $maxTempCelsius = 0
    
    $temperatureData | ForEach-Object {
        # WMI returns temperature in Kelvin x 10, but this class returns Kelvin directly
        # Let's assume the class returns Kelvin (as per your script logic)
        $tempKelvin = $_.Temperature
        $tempCelsius = $tempKelvin - 273.15
        
        if ($tempCelsius -gt $maxTempCelsius) {
            $maxTempCelsius = $tempCelsius
        }
    }
    
    # Calculate Fahrenheit (C * 9/5) + 32
    $tempFahrenheit = ($maxTempCelsius * 1.8) + 32
    
    Write-Host "CPU Temperature (Max):" -ForegroundColor White
    Write-Host "  Celsius: $([math]::Round($maxTempCelsius, 1))°C" -ForegroundColor Green
    Write-Host "  Fahrenheit: $([math]::Round($tempFahrenheit, 1))°F" -ForegroundColor Green
}