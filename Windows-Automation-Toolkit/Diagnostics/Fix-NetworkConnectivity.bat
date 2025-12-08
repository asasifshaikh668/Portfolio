@echo off

echo Diagnosing Network Issues...

REM Flush DNS cache
ipconfig /flushdns

REM Release and renew IP address
ipconfig /release
ipconfig /renew

REM Reset Winsock catalog
netsh winsock reset

REM Reset TCP/IP stack
netsh int ip reset

REM Reset network configuration
netcfg -d

REM Restart network services
netsh interface ipv4 reset
netsh interface ipv6 reset
netsh interface tcp reset
netsh interface teredo reset

REM Check external connectivity post-reset
echo.
echo Attempting to ping Google DNS (8.8.8.8)...
ping 8.8.8.8 -n 4

echo Network Diagnostics Complete.
pause
