
# 🧹 Cleanup-TempFiles
## Script: Cleanup-TempFiles.ps1 & Cleanup-TempFiles.bat

| Section | Content |
|---|---|
| Script Title | Cleanup-TempFiles.ps1 & Cleanup-TempFiles.bat |
| Description | Cleanup-TempFiles is a dual-script utility (PowerShell & Batch) designed to reclaim disk space and resolve common Windows update errors. It aggressively targets temporary user data, system logs, and the Windows Update download cache. |
| Features | ***Temp File Removal:** Deletes data from ```%TEMP%``` and ```C:\Windows\Temp``` to free up space. ***Windows Update Reset:** Stops the ```wuauserv``` service and flushes the ```SoftwareDistribution``` cache to fix stuck updates. ***Recycle Bin Purge:** Forces the Recycle Bin to empty immediately.  ***Browser Cache:** Clears standard browser temporary files located in system temp folders. |
| Usage Instructions | **⚠️ Administrator Privileges Required** > These scripts must be run as ```Administrator``` to access system folders like ```C:\Windows\Temp``` and ```SoftwareDistribution```. **⚠️ Disclaimer:** This script permanently deletes files. While it targets temporary data, always ensure you have saved your work before running.|

********************************************************************************************************

# 🔄 Get-PendingUpdates

## 🗒️ Description
A diagnostic script that audits the Windows Update Agent to provide a clear status report. It lists updates that are currently waiting to be installed and reviews the update history to flag recent installation failures.

## 🚀 Key Features
* **Native Audit**: Uses the built-in `Microsoft.Update.Session` COM object, requiring no third-party modules.
* **Dual Reporting**: Shows both *future* work (Pending) and *past* issues (Failed).
* **Size Estimates**: Calculates download size for pending patches.

## 💻 Usage
1.  Open PowerShell as **Administrator**.
2.  Run the script:
    ```powershell
    .\Get-PendingUpdates.ps1
    ```
3.  Review the output for any Critical updates or recurring failures.

********************************************************************************************************
