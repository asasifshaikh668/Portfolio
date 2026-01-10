# Requires Run as Administrator
Write-Host "🔍 contacting Windows Update Server (this may take a minute)..." -ForegroundColor Cyan

try {
    # Initialize Windows Update Session
    $Session = New-Object -ComObject Microsoft.Update.Session
    $Searcher = $Session.CreateUpdateSearcher()

    # 1. Find Pending Updates (Not Installed & Not Hidden)
    $Search = $Searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
    
    $PendingList = @()
    foreach ($Update in $Search.Updates) {
        $PendingList += [PSCustomObject]@{
            Type        = "Pending Install"
            Title       = $Update.Title
            KB          = $($Update.KBArticleIDs) -join ','
            SizeMB      = [math]::Round($Update.MaxDownloadSize / 1MB, 2)
            IsCritical  = $Update.AutoSelectOnWebSites
        }
    }

    # 2. Find Recently Failed Updates (Last 30 entries)
    # ResultCode: 2=Succeeded, 4=Failed, 5=Aborted
    $History = $Searcher.QueryHistory(0, 30) | Where-Object { $_.ResultCode -eq 4 }
    
    $FailedList = @()
    foreach ($H in $History) {
        $FailedList += [PSCustomObject]@{
            Type   = "Previous Failure"
            Title  = $H.Title
            KB     = "N/A" # History object structure differs slightly
            SizeMB = "N/A"
            Date   = $H.Date
        }
    }

    # Display Results
    if ($PendingList) {
        Write-Host "`n📥 PENDING UPDATES:" -ForegroundColor Yellow
        $PendingList | Format-Table -AutoSize
    } else {
        Write-Host "`n✅ System is up to date." -ForegroundColor Green
    }

    if ($FailedList) {
        Write-Host "`n❌ RECENT INSTALLATION FAILURES:" -ForegroundColor Red
        $FailedList | Select-Object Date, Title | Format-Table -AutoSize
    }

} catch {
    Write-Error "Failed to query Windows Update Agent. Ensure you are running as Administrator."
}