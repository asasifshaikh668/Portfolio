@echo off
:: Move to the script's directory (ensures relative paths like 'Images' work)
cd /d "%~dp0"

:: Launch the PowerShell script with the Bypass flag
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "IT-Support-Automation-Suite.ps1"