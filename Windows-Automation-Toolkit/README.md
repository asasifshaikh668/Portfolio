# 🛠️ Windows-Automation-Toolkit
A collection of PowerShell and Batch scripts developed to streamline Windows system administration and enhance Technical Support efficiency. Includes tools for diagnostics, hardware auditing, system component retrieval, and automated task execution.

# 📂 Repository Structure

**1. 🔍 Diagnostic/**

Contains scripts focused on system analysis, error logging, and troubleshooting.

    -Includes 15+ scripts for verifying system health, network connectivity, and resource usage.


**2. 🛡️ System-Security-Utilities/**

Contains tools for active maintenance, security auditing, and configuration management.

    -Cleanup-TempFiles (Batch/PS): Automates disk cleanup and cache removal.

    -Install-Software-Silent (PowerShell): Deploys support tools without user interaction.

    -Get-PendingUpdates (PowerShell): Audits Windows Update status.

    -Check-LocalAdminAccess (PowerShell): Lists members of the local Administrators group.

    -Disable-UAC (Batch/PS): Adjusts User Account Control settings.


**3. 🔋 Hardware-Health/**

Contains dedicated scripts for monitoring physical component status and longevity.

    -Get-CPUTemperature (PowerShell): Monitors thermal performance to prevent overheating.

    -check-DiskHealth (PowerShell): Reports drive wear levels and SMART status.

    -Get-BatteryHealth (PowerShell): Generates detailed battery health and capacity reports.
