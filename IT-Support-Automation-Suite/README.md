# IT-Support-Automation-Suite

A professional PowerShell WinForms diagnostic and automation tool designed for Helpdesk, Desktop, and Technical Support Engineers.
A comprehensive, GUI-based PowerShell utility designed to standardize troubleshooting, accelerate resolution times, and empower IT professionals.

# Introduction

The IT-Support-Automation-Suite is a standalone application built entirely in PowerShell, utilizing Windows Forms (WinForms) to provide a rich Graphical User Interface (GUI).
It consolidates over 50+ critical administrative tasks—ranging from deep system diagnostics and network repair to hardware health monitoring and Intune compliance—into a single, portable interface.
It requires no installation, making it the perfect "Swiss Army Knife" for IT Helpdesk/Desktop/Technical Support Engineers, System Administrators, and Network Engineers.

# Why This Toolkit Exists

In complex enterprise environments, IT teams often face three core challenges:

  - **Inconsistency**: different analysts use different methods to fix the same "slow PC" or "network error" ticket.
  - **Inefficiency**: valuable time is wasted navigating deep Windows settings menus or memorizing complex command-line arguments.
  - **Accessibility**: powerful fixes (like resetting TCP/IP stacks or analyzing Event Logs) are often locked behind a "skill gap," accessible only to senior engineers.

This toolkit solves these problems by wrapping complex PowerShell logic into simple, one-click buttons.

# Business Value

This project delivers immediate value to IT Operations by focusing on:
- **Drastically Reduced MTTR (Mean Time To Resolve)**: Complex diagnostic sequences that typically take 15–20 minutes manually are executed in seconds.
- **Operational Safety**: Features like "Soft Browser Reset" and read-only "Network Monitoring" allow aggressive troubleshooting without the risk of data loss.
- **Security Compliance**: Built-in "Run as Admin" checks ensure that sensitive tasks are only performed with elevated privileges, while standard tasks remain accessible.
- **Standardization**: Ensures that every "Network Reset" or "Teams Cache Clear" is performed identically across the entire department, reducing ticket re-opens.

# ✨ Key Features

- **Diagnostic Hub**: Automates Windows Updates, clears temp files, and fixes common OS corruption (SFC/DISM).
- **Network Center**: Tools for flushing DNS, resetting Winsock, analyzing Wi-Fi profiles, and monitoring LibreNMS status.
- **Log Analyzer**: Instantly parses Event Viewer for crashes, Blue Screens (BSOD), and user login history without digging through the Eventvwr GUI.
- **Hardware Health**: Proactive monitoring of SSD wear levels, battery health cycles, and CPU thermal throttling.

# How to Run

This toolkit is a single .ps1 script. No .exe compilation or installation is required.

- Download the Launch-Windows-Automation-Toolkit.ps1 file.
- Right-click the file and select "Run with PowerShell".
- (Optional) If prompted about Execution Policy, run this command in PowerShell first:

PowerShell
```Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process```

*****************************************************************************************************************************************************************************

# 🔧 Detailed Feature Breakdown

# TAB 1: DIAGNOSTIC (The "First Response" Tools)

| Button Name | What It Does (Technical) | Business Value |
| :--- | :--- | :--- |
| **Cleanup Temp Files** | Deletes %TEMP%, C:\Windows\Temp, and clears SoftwareDistribution. | Frees space and fixes 80% of Windows Update issues. |
| **Windows Update** | Uses the COM Object (Microsoft.Update.Session) to force-scan, accept EULAs programmatically, and install updates. | Bypasses the Windows Settings UI glitches; forces updates even when the UI says "Error". |
| **Check Disk Space (C:)** | RQueries WMI for C: drive space. Alerts if below 15GB. | Prevents system crashes due to full disks. |
| **Reset Print Spooler** | Stops Spooler service, deletes files in System32\spool\PRINTERS, restarts service. | This one fix for "Stuck Print Jobs" without needing a reboot. |
| **System Uptime** | GCalculates time since LastBootUpTime. | Catch users who say "I restarted" when they actually just closed the lid (Fast Startup). |
| **System Information** | Pulls Serial Number, TPM Version, RAM, and Azure AD Join status. | fast asset verification during a call. |
| **Installed Software** | Lists software from Registry (Uninstall keys). Includes a search filter. | Allows uninstallation of apps that don't appear in the standard "Add/Remove Programs" list. |
| **BitLocker Status** | Runs Get-BitLockerVolume. (Requires Admin). | Verifies if the drive is encrypted and lists Key Protectors (TPM/PIN). |
| **Monitor Peaks (60s)** | Logs CPU & RAM usage every second for 1 minute to CSV. | Catch intermittent "freezing" that happens during video calls. |
| **Pending Reboot** | Checks Registry keys (ComponentBasedServicing, WindowsUpdate). | Tells you if a reboot is actually required before you waste time troubleshooting. |
| **Local Admin Status** | Lists all local users in the "Administrators" group. | Security Audit: Finds unauthorized local admin accounts. |
| **Audio Troubleshooter** | Restarts Audiosrv and launches built-in MSDT diagnostic packs. | Fixes "No Audio Device Installed" errors. |
| **Windows Activation** | Checks SoftwareLicensingProduct WMI class. | Verifies if Windows is genuine and activated. |
| **Teams Cache Delete** | Kills Teams process and wipes the AppData\...\MSTeams folder. | Fixes login loops, missing profile pictures, and old chat history issues. |
| **Outlook Repair** | Offers switches: /safe, /cleanviews, /resetnavpane. | Fixes Outlook not opening or "Processing..." hangs. |
| **Defender Full Scan** | Triggers Start-MpScan -ScanType FullScan. | Initiates a virus scan without navigating the security center. |
| **Defender Logs** | Opens C:\ProgramData\...\Support. | Quick access to MPLog.log for malware analysis. |
| **Intune Device Mgr** | (Sub-Menu) Checks App Cache, Forces Agent Sync, Restarts IME Service. | Resolves "Not Compliant" issues and stuck app installations. |
| **Browser Repair** | (Sub-Menu) Soft Reset (Renames Profile) for Edge/Chrome/Firefox. | Fixes browser corruption while keeping a backup of user Bookmarks. |


# TAB 2: HARDWARE HEALTH

| Button Name | What It Does (Technical) | Business Value |
| :--- | :--- | :--- |
| **SSD Health** | Queries Storage Reliability Counters (Wear Level). | Predicts hard drive failure before data is lost. |
| **Battery Health** | Generates XML report comparing Design Capacity vs Full Charge. | Tells you if a laptop battery needs replacement (e.g., 50% health). |
| **CPU Temp** | WMI ThermalZone query. | Checks if a PC is overheating (fan failure). |
| **Thermal Log** | Logs temperature over 60 seconds. | Diagnoses overheating under load. |


# TAB 3: NETWORK (Connectivity Tools)

| Button Name | What It Does (Technical) | Business Value |
| :--- | :--- | :--- |
| **Test Remote Conn.** | Performs a Test-Connection (Ping) and Test-NetConnection (Port Scan) on 80/443/3389. | Tells you if a server is down OR if a firewall port is blocked. |
| **Fix Network Conn.** | Runs ipconfig /flushdns, netsh winsock reset, netcfg -d. | The ultimate "Internet Repair" button. Fixes corruption in the network stack. |
| **Repair Winsock/TCP** | Targeted reset of Winsock Catalog. | Specific fix for VPN clients failing to connect. |
| **Get Public IP** | Queries api.ipify.org. | Checks if the user is on the corporate gateway or home ISP. |
| **Adapter Details** | Lists Link Speed, MAC Address, and Status. | Identifies if a cable is unplugged or negotiating at wrong speeds (e.g., 100mbps vs 1Gbps). |
| **Saved Wi-Fi Pass** | Extracts Key Content from netsh wlan show profiles. | Recovers forgotten Wi-Fi passwords for the user. |
| **Trace Route** | Runs Test-NetConnection -TraceRoute. | Identifies exactly where packets are dropping (ISP vs Internal). |
| **DNS Lookup** | Runs Resolve-DnsName. | Verifies if DNS is resolving internal server names correctly. |
| **Check SSL Expiry** | Scrapes the SSL Certificate expiration date from a URL. | Quick check to see if a website "down" issue is actually an expired cert. |
| **Flush DNS** | Runs Clear-DnsClientCache. | Quick fix for users who can't reach a website that moved IPs recently. |
| **IP Release/Renew** | Standard DHCP cycle. | Refreshes local IP lease. |
| **IP Config /All** | GUI view of ipconfig /all. | Full network detail audit. |
| **Set IP Auto** | Sets interface to DHCP. | Fixes static IP misconfigurations left by users. |
| **Reset All Adapters** | Loops through ALL adapters and sets to DHCP. | "Factory Reset" for network settings. |
| **Hosts File Mgr** | Reads/Writes to drivers\etc\hosts. | Safe way to edit hosts file for dev testing or blocking sites. |


# TAB 4: LOGS (Event Viewer Analysis)

| Button Name | What It Does (Technical) | Business Value |
| :--- | :--- | :--- |
| **Event Log Errors** | Fetches Level 1 & 2 events from System/App logs. | Shows why an app crashed without digging through Event Viewer. |
| **Boot/Shutdown** | Filters Event IDs 6005, 6006, 6008, 1074. | Proves if a computer crashed (Dirty Shutdown) or was restarted by a user. |
| **User Login History** | Filters Security Event 4624/4625. | Audits who logged into the PC and when. |
| **Search Logs** | Keywords search across logs. | Find specific error codes or app names fast. |
| **Export Errors CSV** | Dumps logs to a file. | Creates a file to send to L3/Engineering for analysis. |
| **Account Lockouts** | Filters Event 4740. | Shows source of account lockouts (which PC is locking the account). |
| **App Crash Report** | Filters Event 1000. | Lists every application crash timestamp. |
| **BSOD History** | Filters Event 1001 (BugCheck). | Shows if the Blue Screen of Death has occurred recently. |
| **Win Update History** | Filters WindowsUpdateClient events. | History of successful/failed patch installations. |
| **Printer Errors** | Filters PrintSpooler events. | Diagnoses why a printer failed to print. |

************************************************************************************************************************************************************************************

# ⚠️ Disclaimer

This tool is provided "as is" without warranty of any kind. While every effort has been made to ensure safety (e.g., using "Soft Resets" instead of deletions), always test in a non-production environment before deploying to critical systems.

*************************************************************************************************************************************************************************************

# 👤 Author
**Mohammed Asif Shaikh** _Senior IT Helpdesk Analyst | Automation Enthusiast_

If you find this tool helpful, please give it a ⭐️ on GitHub!




