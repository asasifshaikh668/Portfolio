# --- 1. TIME ZONE MAPPING (CORRECTED) ---
$TimeZoneMap = @{
    "1" = @{Name='India (IST)'; ID='India Standard Time'}
    "2" = @{Name='London (GMT/BST)'; ID='GMT Standard Time'}
    "3" = @{Name='New York (EST/EDT)'; ID='Eastern Standard Time'}
    "4" = @{Name='Australia (Sydney)'; ID='AUS Eastern Standard Time'}
    "5" = @{Name='Hong Kong'; ID='China Standard Time'}
    "6" = @{Name='China'; ID='China Standard Time'}
    "7" = @{Name='Japan'; ID='Tokyo Standard Time'}
    "8" = @{Name='Germany'; ID='W. Europe Standard Time'}
    "9" = @{Name='France'; ID='W. Europe Standard Time'}
    "10" = @{Name='Dubai (GST)'; ID='Arabian Standard Time'}
    # ... all other entries corrected
}

function Show-Menu {
    param($Title, $Map)
    
    Write-Host "`n================================================"
    Write-Host $Title -ForegroundColor Yellow
    Write-Host "================================================"
    
    $Map.Keys | Sort-Object { [int]$_ } | ForEach-Object {
        Write-Host "[$_] $($Map.$_.Name)" -ForegroundColor Cyan # Quoted output string
    }
    Write-Host "[0] Exit" -ForegroundColor Red
    Write-Host "------------------------------------------------"
    return (Read-Host -Prompt 'Enter selection') # Quoted prompt
}

#--- 2. REST OF SCRIPT LOGIC (CORRECTED) ---

# --- SOURCE TIME ZONE SELECTION (Your code ends here) ---
do {
    $SourceChoice = Show-Menu 'SELECT SOURCE TIME ZONE' $TimeZoneMap
    if ($SourceChoice -eq '0') { return }
    $SourceZone = $TimeZoneMap[$SourceChoice]
} while (-not $SourceZone)

# --- DESTINATION TIME ZONE SELECTION (Add this block next) ---
do {
    # Calls the same menu function but stores the result in Target-specific variables
    $TargetChoice = Show-Menu 'SELECT DESTINATION TIME ZONE' $TimeZoneMap
    if ($TargetChoice -eq '0') { return }
    $TargetZone = $TimeZoneMap[$TargetChoice] # Stores the destination zone details
} while (-not $TargetZone)

# --- TIME INPUT AND CONVERSION (This block follows the Target loop) ---
# $InputTimeStr = Read-Host -Prompt "Enter the time..."
# ... (Conversion logic using $SourceZone and $TargetZone) ...

# --- 3. GET TIME INPUT ---
$InputTimeStr = Read-Host -Prompt "Enter the time in $($SourceZone.Name) (e.g., 8:00 PM or 20:00 IST)"


# --- 4. PARSE, CONVERT, AND DISPLAY (COMPLETE CODE) ---
try {
    # 4a. Combine today's date with the user's time input and assign the source timezone
    # Get-Date is used to get the current date [6](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-7.5).
    $Today = Get-Date -Format "MM/dd/yyyy"
    $FullDateTimeStr = "$Today $InputTimeStr"
    
    # 4b. Define Time Zone Info Objects (Required for the conversion method)
    # This relies on the ID property retrieved from the $TimeZoneMap.
    $SourceZoneInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($SourceZone.ID)
    $TargetZoneInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($TargetZone.ID)

    # 4c. Create a DateTime object using the user's input string
    $DateTimeToConvert = [DateTime]::Parse($FullDateTimeStr)
    
    # 4d. Perform the Conversion
    # [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId is used to accurately convert the time between the two specified zones [3](https://www.craigforrester.com/posts/convert-times-between-time-zones-with-powershell/).
    $ConvertedTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($DateTimeToConvert, $SourceZoneInfo.Id, $TargetZoneInfo.Id)

    # --- 5. Display Results ---
    Write-Host "`n================================================"
    Write-Host " Meeting Time Conversion Result" -ForegroundColor Yellow
    Write-Host "================================================"
    Write-Host "Source Time Zone: $($SourceZone.Name)" -ForegroundColor Cyan
    Write-Host "Target Time Zone: $($TargetZone.Name)" -ForegroundColor Cyan
    Write-Host "------------------------------------------------"
    Write-Host "Input Time ($($InputTimeStr)): $($DateTimeToConvert.ToString('hh:mm tt, ddd MMM dd'))" -ForegroundColor Green
    Write-Host "CONVERTED TIME: $($ConvertedTime.ToString('hh:mm tt, ddd MMM dd'))" -ForegroundColor Green
    Write-Host "================================================"
    
}
catch {
    # This block executes if the user enters a bad time format (e.g., "abc") or if a TimeZone ID is corrupt.
    Write-Host "`nERROR: Conversion failed. Check time format (e.g., 8 PM or 20:00) and try again.`n" -ForegroundColor Red
    $_.Exception.Message # Output the specific error message for troubleshooting
}

# --- 6. FINAL PAUSE (Ensures window stays open) ---
Read-Host -Prompt "Press Enter to exit" 
# Note: If you have an outer loop, this is where you replace Read-Host with the 'Continue/Exit' prompt.



# --- 5. FINAL PAUSE / LOOP CONTROL ---
Read-Host -Prompt "Press Enter to exit"
# OR, if you implemented the outer loop:
# $ContinueChoice = Read-Host -Prompt "Conversion complete. Run another conversion? (Y/N)"
