# GlobalTimeConverter


| Section | Content |
|---|---|
| Script Title | GlobalTimeConverter.ps1 |
| Description | This script is a powerful utility for scheduling cross-time zone meetings and calculating time differences accurately. It prompts the user for a specific time (e.g., 4:30 PM) and instantly converts it between two predefined time zones (e.g., IST and UK time).|
| Key Features & Commands | **Complex Conversion:** Uses the .NET method [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId for reliable conversion, including automatic handling of Daylight Saving Time (DST). **Interactive Input:** Uses Read-Host to prompt the user for the exact time to be converted. **Predefined Zones:** Hardcoded with common support time zones for instant use, but easily customizable by the engineer. **Formatted Output:** Provides clean, color-coded results in the 12-hour format (hh:mm tt). |
| Technical Support Value & Impact | **Scheduling Efficiency:** Eliminates all manual calculation errors when booking appointments with international clients or remote teams. **Professionalism:** Demonstrates advanced PowerShell capability beyond simple command execution. |
| How to Run | Run the wrapper script: .\Run-TimezoneConverter.bat. No Admin rights required. The script will prompt for the time input. |
