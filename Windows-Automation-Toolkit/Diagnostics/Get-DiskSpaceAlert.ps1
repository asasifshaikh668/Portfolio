function Get-DiskSpaceAlert {
    param(
        [string]$Drive = "C:",
        [int]$ThresholdGB = 10
    )

    $Disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$Drive'"
    $FreeSpaceGB = [math]::Round($Disk.FreeSpace / 1GB, 2)

    if ($FreeSpaceGB -lt $ThresholdGB) {
        Write-Warning "Low Disk Space! $Drive only has $FreeSpaceGB GB remaining."
        # Log to Application Event Log
        Write-EventLog -LogName Application -Source "DiskMonitor" -EntryType Warning `
            -EventID 1001 -Message "Drive $Drive is below $ThresholdGB GB threshold. Current: $FreeSpaceGB GB."
    } else {
        Write-Host "$Drive has sufficient space: $FreeSpaceGB GB." -ForegroundColor Green
    }
}

Get-DiskSpaceAlert -ThresholdGB 15