# Run as Administrator (Required for Get-StorageReliabilityCounter)

# --- 1. Find the Physical Disk for C: Drive ---
# Get the partition object for drive C:, then pipe it to Get-Disk to get the associated PhysicalDisk object.
$C_Drive = Get-Partition -DriveLetter C | Get-Disk
$disk = $C_Drive | Get-PhysicalDisk -ErrorAction Stop # Use Get-PhysicalDisk to ensure the correct object type for the next step

if (-not $disk) {
    Write-Host "ERROR: C: drive physical disk could not be identified or accessed." -ForegroundColor Red
    return
}

# --- 2. Retrieve Reliability Data (For C: Drive's physical disk) ---
$reliability = Get-StorageReliabilityCounter -PhysicalDisk $disk -ErrorAction SilentlyContinue

Write-Host "`n--- 💡 C: Drive Health Report ---`n" -ForegroundColor Yellow

# --- 3. Output Disk Info ---
Write-Host "Disk: $($disk.FriendlyName)" -ForegroundColor White
Write-Host "Model: $($disk.Model)" -ForegroundColor White
Write-Host "Media Type: $($disk.MediaType)" -ForegroundColor White

# --- 4. Output Reliability Data ---
if ($reliability) {
    Write-Host "Temperature: $($reliability.Temperature) °C" -ForegroundColor Green
    
    # Wear level calculation
    if ($reliability.Wear -ne $null) {
        $wear = $reliability.Wear
        $health = 100 - $wear
        Write-Host "Wear Level: $wear%" -ForegroundColor Cyan
        Write-Host "Disk Health: $health%" -ForegroundColor Green
        
        if ($health -lt 80) {
             Write-Host "⚠️ WARNING: Drive health is below 80%. Consider replacement." -ForegroundColor Red
        }
    } else {
        Write-Host "Wear Level: Not reported (Common for some HDDs/Virtual Disks)." -ForegroundColor White
    }
} else {
    Write-Host "Reliability Data: Not available or access denied." -ForegroundColor Red
}

# --- 5. Output Volume Information for C: ---
# Skip the complex loop and target the C: volume directly
$C_Volume = Get-Volume -DriveLetter C

if ($C_Volume -and $C_Volume.FileSystem -ne 'RAW') {
    $usedGB = [math]::Round(($C_Volume.Size - $C_Volume.SizeRemaining) / 1GB, 2)
    $freeGB = [math]::Round($C_Volume.SizeRemaining / 1GB, 2)
    $totalGB = [math]::Round($C_Volume.Size / 1GB, 2)

    Write-Host "`nVolume: $($C_Volume.DriveLetter): ($($C_Volume.FileSystem))" -ForegroundColor DarkCyan
    Write-Host "  Total Space: $totalGB GB"
    Write-Host "  Free Space: $freeGB GB"
    Write-Host "  Used Space: $usedGB GB"
}

Write-Host `n-----------------------------`n