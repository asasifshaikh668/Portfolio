# Script: Invoke-SystemDiagnostics.bat

This script is designed to rapidly diagnose and proactively address common causes of Windows system slowness in a single execution. It significantly reduces the time spent manually collecting performance data and running standard maintenance tasks during a support ticket lifecycle.

| Category | Command/Function | Purpose & Value |
| :--- | :--- | :--- |
| Initial Assessment | wmic cpu get load percentage, wmic os get physical memory, wmic logical disk get size,freespace,caption | ollects baseline metrics (CPU, RAM, Disk space/free space) to immediately identify hardware bottlenecks (e.g., low RAM or full disk) before running repairs. |
| Running Processes | Tasklist | Identifies resource-intensive, unauthorized, or runaway processes that may be consuming excessive resources, leading to slowness. |
| Startup Optimization | wmic startup get caption,command | Lists applications configured to run at startup, enabling identification of unnecessary programs slowing down boot time and initial login. |
| Cleanup | Cleanmgr /d C:, Rmdir /s /q %temp% | Executes built-in Disk Cleanup utility and removes temporary files (%temp%), reclaiming disk space and improving read/write speeds. |
| System Integrity Checks | sfc /scannow, Dism /online /cleanup-image /restorehealth | Runs System File Checker (SFC) and Deployment Image Servicing and Management (DISM) to repair corrupted system files and restore the health of the Windows component store. |
| File System Health | Defrag C: /U /V, Chkdsk /f/r | Optimizes file structure (defragmentation) for mechanical drives and checks/repairs the file system structure (CHKDSK), addressing disk-related performance degradation. |


**Technical Support Value & Impact**

* **Time Reduction:** Automates over 10 separate manual commands, saving an estimated **15–20 minutes** per troubleshooting session.
* **Proactive Repair:** Unlike manual diagnosis, this script immediately initiates repair steps (sfc, Dism, Cleanmgr), often resolving performance issues without further technician interaction.
* **Standardization:** Ensures every system slowness ticket is diagnosed and treated using the same **best-practice methodology**.

----------------------------------------------------------
**How to Run**

1. Download the script to the target Windows machine.
2. Open Command Prompt as an Administrator.
3. Execute the script:
   
   ```.\Invoke-SystemDiagnostics.bat```
   
*Note: Script execution may take several minutes due to comprehensive tasks like CHKDSK and SFC.*


******************************************************************************************************


# Script: Repair-AudioTroubleshooter.bat

| Section | Content |
|---|---|
| Script Title | Repair-AudioTroubleshooter.bat |
| Description | This script provides a one-click solution for diagnosing and fixing common Windows audio playback and recording problems, often eliminating the need for manual service restarts and driver re-enabling. It integrates built-in Windows troubleshooters with backend service controls and device manipulation. |
| Key Features & Commands | * Built-in Troubleshooting: Executes both the Audio Playback and Audio Recording diagnostics (msdt.exe). * Service Reset: Stops and restarts the core Windows Audio service (Audiosrv). * Device Cycling: Uses PowerShell to identify audio devices in an 'Error' state and automatically disables and re-enables them (a simulated hardware reset). * DLL Repair (Added): Re-registers audiosrv.dll and audiocore.dll to ensure core system files are correctly registered. |
| Technical Support Value & Impact | * Time Reduction: Automates a 5-step manual process into a single execution, saving 5-10 minutes per ticket. * High Success Rate: Addresses the top three causes of audio failure (corrupt services, disabled devices, and software errors) simultaneously. |
| How to Run | Prerequisite: Must be run with Administrator privileges. Execute the script: .\Repair-AudioTroubleshooter.bat |

******************************************************************************************************


# Script: Fix-NetworkConnectivity.bat

| Section | Content |
|---|---|
| Script Title | Fix-NetworkConnectivity.bat |
| Description | This script executes a full, low-level reset and repair of the Windows networking stack (TCP/IP, Winsock, IPv4/v6) to resolve common connectivity issues such as limited connectivity, slow speeds, inability to resolve DNS, and DHCP failures. It is a crucial one-click tool for Level 1/2 network troubleshooting. |
| Key Features & Commands | ***IP Reset:** Runs ipconfig /release and ipconfig /renew to get a fresh DHCP lease. ***DNS Flush:** Clears the local DNS cache (ipconfig /flushdns). ***Winsock/IP Reset:** Resets the Winsock Catalog and TCP/IP stack (netsh winsock reset, netsh int ip reset), which resolves many corrupted protocol issues. ***Full Network Interface Reset:** Runs multiple netsh commands (ipv4 reset, ipv6 reset, tcp reset) to reset all related network configurations. |
| Technical Support Value & Impact | ***Efficiency:** Replaces a long, error-prone list of manual commands, standardizing the network repair process. ***Speed:** Quickly addresses the core of most network issues caused by temporary software corruption. |
| How to Run | Prerequisite: Must be run with Administrator privileges. Execute the script: .\Fix-NetworkConnectivity.bat |

******************************************************************************************************


# Script: Get-GlobalTime.bat

| Section | Content |
|---|---|
| Script Title | Get-GlobalTime.bat |
| Description | This interactive command-line utility provides instant, accurate time and date information for 14 major global cities and regions. It is an essential tool for Technical Support Engineers who need to coordinate with international clients, schedule calls, or determine remote user availability across various time zones. |
| Key Features & Commands | ***Interactive Menu:** Uses Batch scripting's GOTO and SET /P to provide a simple, numbered interface. ***Accurate Time Calculation:** Leverages the PowerShell Get-Date -TimeZone cmdlet to ensure correct time calculation, including automatic handling of Daylight Saving Time (DST) for each region. ***Hybrid Scripting:** Demonstrates proficiency in embedding advanced PowerShell functionality within traditional Batch files. |
| Technical Support Value & Impact | ***Coordination:** Eliminates reliance on web browsers for simple time checks, improving workflow during urgent international support operations. ***Professionalism:** Ensures communication with global clients is always timely and accurate. |
| How to Run | Execute the script: .\Get-GlobalTime.bat. No Admin rights required. Follow the on-screen menu instructions. |

******************************************************************************************************


# Script: GlobalTimeConverter.ps1

| Section | Content |
|---|---|
| Script Title | GlobalTimeConverter.ps1 |
| Description | This script is a powerful utility for scheduling cross-time zone meetings and calculating time differences accurately. It prompts the user for a specific time (e.g., 4:30 PM) and instantly converts it between two predefined time zones (e.g., IST and UK time). |
| Key Features & Commands | ***Complex Conversion:** Uses the .NET method [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId for reliable conversion, including automatic handling of Daylight Saving Time (DST). ***Interactive Input:** Uses Read-Host to prompt the user for the exact time to be converted. ***Predefined Zones:** Hardcoded with common support time zones for instant use, but easily customizable by the engineer. ***Formatted Output:** Provides clean, color-coded results in the 12-hour format (hh:mm tt). |
| Technical Support Value & Impact | ***Scheduling Efficiency:** Eliminates all manual calculation errors when booking appointments with international clients or remote teams. ***Professionalism:** Demonstrates advanced PowerShell capability beyond simple command execution. |
| How to Run | Run the wrapper script: **.\GlobalTimeConverter.bat.** No Admin rights required. The script will prompt for the time input. |


******************************************************************************************************
