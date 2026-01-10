function Get-DetailedRestorePoints {
    Write-Host "🔍 SYSTEM RESTORE POINTS (User View)" -ForegroundColor Cyan
    Write-Host "------------------------------------" -ForegroundColor Gray
    
    # 1. Get Restore Points with Creation Date
    $RPoints = Get-ComputerRestorePoint
    if ($RPoints) {
        $RPoints | Select-Object SequenceNumber, @{N='Created On';E={$_.CreationTime}}, Description, RestorePointType | 
        Format-Table -AutoSize
    } else {
        Write-Warning "No System Restore points found."
    }

    Write-Host "`n📂 SHADOW COPY STORAGE LOCATIONS (System View)" -ForegroundColor Cyan
    Write-Host "---------------------------------------------" -ForegroundColor Gray
    
    # 2. Get Underlying Shadow Copy Paths (Location)
    # This shows where the data actually lives on the disk structure.
    Get-CimInstance -ClassName Win32_ShadowCopy | Select-Object `
        @{N='Shadow Copy ID';E={$_.ID}},
        @{N='Created On';E={$_.InstallDate}},
        @{N='Original Volume';E={$_.VolumeName}},
        @{N='Storage Location (Device Path)';E={$_.DeviceObject}} | 
        Format-Table -AutoSize
}

Get-DetailedRestorePoints