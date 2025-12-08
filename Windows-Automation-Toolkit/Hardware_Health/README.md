# Hardware Health and Preventative Maintenance Scripts

This module contains a set of PowerShell scripts designed for proactive technical support. These tools retrieve internal hardware metrics that are often leading indicators of system stability issues or impending component failure.

----------------------------------------------------------------------------------------------------------

# Get-CPUTemperature.ps1 (CPU Temperature Check)

| Feature | Description |
|---|---|
| Script Title | Get-CPUTemperature.ps1 |
| Purpose | Queries the system's **WMI (Windows Management Instrumentation)** to estimate the CPU's current temperature. High temperatures indicate failing fans, dust buildup, or poor thermal paste application. |
| Core Commands | Get-WmiObject -Class Win32_PerfFormattedData_Counters_ThermalZoneInformation|
| Output | Provides the CPU temperature in both **Celsius (°C)** and **Fahrenheit (°F)**. Includes error handling for systems that do not expose this data. |
| Technical Support Value & Impact | Quick validation for performance issues and spontaneous shutdowns linked to thermal throttling. |
| How to Run | Run the script: .\Get-CPUTemperature.ps1. Admin rights is required to run the cript.|

----------------------------------------------------------------------------------------------------------

# Check-DiskHealth.ps1 (Drive Wear and Status)

| Feature | Description |
|---|---|
| Script Title | Check-DiskHealth.ps1 |
| Purpose | Performs a detailed health check for the physical disk hosting the C: drive. It uses **S.M.A.R.T.** data to retrieve component longevity metrics. SSD health decreases due to wear from write operations over time 5, making this essential for predictive maintenance.|
| Core Commands | Get-Partition, Get-Disk, Get-PhysicalDisk, Get-StorageReliabilityCounter|
| Output | Reports the Disk Model, Temperature, **Wear Level (%)**, calculated Disk Health (100% - Wear), and the Volume capacity/usage of the C: drive. |
| Technical Support Value & Impact | Allows support staff to proactively recommend a replacement when disk health drops below a warning threshold (e.g., 80%), preventing data loss. |
| How to Run | Run the script: .\Check-DiskHealth.ps1. Admin rights is required to run the cript.|

----------------------------------------------------------------------------------------------------------

# Get-BatteryHealth.ps1 (Detailed Battery Health Report)

| Feature | Description |
|---|---|
| Script Title | Get-BatteryHealth.ps1 |
| Wear Level Calculation | Calculates the **Health Percentage** by comparing the current Full Charge Capacity to the Design Capacity using WMI data. |
| Current Status | Displays the immediate status, including Estimated Charge Remaining (%) and Battery Status (Charging/Discharging). |
| HTML Report Integration | Automatically executes the powercfg /batteryreport command and saves the full historical report to the user's %TEMP% directory. |
| Technical Support Value & Impact | Engineers can quickly quantify battery degradation to justify a replacement, instead of relying on subjective user reports. |
| How to Run | Run the script: .\Get-BatteryHealth.ps1. Admin rights is required to run the cript. |
