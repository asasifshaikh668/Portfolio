@echo off
echo Cleaning User and System Temp Files...
del /q /f /s "%TEMP%\*"
del /q /f /s "%WINDIR%\Temp\*"

echo Emptying Recycle Bin...
rd /s /q %systemdrive%\$Recycle.bin

echo Cleaning Windows Update Cache...
net stop wuauserv
del /f /s /q "%WINDIR%\SoftwareDistribution\Download\*"
net start wuauserv

echo Cleanup Complete.
pause