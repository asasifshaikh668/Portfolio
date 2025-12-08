# --- Generate Battery Report ---
 
$BatteryReportPath = "$env:C:\Windows\System32\battery-report.html"
 
powercfg /batteryreport /output "$BatteryReportPath"
# --- Generate Energy XML Report ---
 
$xmlPath = "$env:TEMP\energy-report.xml"
 
if (-not (Test-Path $xmlPath)) {
 
    powercfg /energy /output $xmlPath /xml | Out-Null
 
    Start-Sleep -Seconds 5
 
}
# --- Parse Battery Health Details from XML ---
 
[xml]$xml = Get-Content $xmlPath
 
$nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
 
$nsmgr.AddNamespace("ns", "http://schemas.microsoft.com/energy/2007")
$designCap = 0
 
$lastFull = 0
 
$health = 0
 
$wear = 0
$entries = $xml.SelectNodes("//ns:Detail", $nsmgr)
 
foreach ($entry in $entries) {
 
    $name = $entry.SelectSingleNode("ns:Name", $nsmgr).InnerText
 
    $value = $entry.SelectSingleNode("ns:Value", $nsmgr).InnerText
    if ($name -like "*Design Capacity*") {
 
        $designCap = ($value -replace "[^\d]", "") -as [int]
 
    }
 
    elseif ($name -like "*Last Full Charge*") {
 
        $lastFull = ($value -replace "[^\d]", "") -as [int]
 
    }
 
}
if ($designCap -gt 0 -and $lastFull -gt 0) {
 
    $health = [math]::Round(($lastFull / $designCap) * 100, 2)
 
    $wear = 100 - $health
 
}
# --- Display Health Info in Terminal ---
 
Write-Host "`n--- 💡 Battery Health Estimation ---`n"
 
Write-Host "Battery Charge Remaining    : $($battery.EstimatedChargeRemaining)%"
 
Write-Host "battery Design Capacity     : $designCap mWh"
 
Write-Host "battery Last Full Charge    : $lastFull mWh"
 
Write-Host "Battery Health      : $health %"
 
Write-Host "Battery Wear Level  : $wear %"

Write-Host "`n-----------------------------`n"
# --- Create HTML Output ---
 
$html = @"
<html>
<head><title>Battery Health Report</title></head>
<body>
<h1>🔋 Battery Health Report</h1>
<p><b>Design Capacity:</b> $designCap mWh</p>
<p><b>Last Full Charge:</b> $lastFull mWh</p>
<p><b>Battery Health:</b> $health%</p>
<p><b>Battery Wear Level:</b> $wear%</p>
</body>
</html>
 
"@