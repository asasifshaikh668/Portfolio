@echo off

REM *** SCRIPT: Invoke-SystemDiagnostics.bat ***
REM *** Ensures Administrative privileges are used ***

NET SESSION >NUL 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: This script requires Administrator privileges.
    echo Please right-click the file and select "Run as administrator".
    pause
    goto :EOF
)

echo.
echo ===========================================
echo   Starting System Slowness Diagnostics
echo ===========================================
echo.

REM --- 1. DIAGNOSTICS & BASELINE METRICS ---

echo [1/11] Checking CPU Load...
wmic cpu get loadpercentage /value
echo.

echo [2/11] Checking Physical Memory...
wmic os get freephysicalmemory /value
echo.

echo [3/11] Checking Disk Usage (Size and FreeSpace)...
wmic logicaldisk get size,freespace,caption
echo.

echo [4/11] Checking Running Processes (Tasklist)...
tasklist
echo.

echo [5/11] Checking Startup Programs...
wmic startup get caption,command
echo.

REM --- 2. SYSTEM INTEGRITY AND CLEANUP ---

echo [6/11] Running Windows System File Checker (SFC)...
sfc /scannow
echo.

echo [7/11] Running Deployment Image Servicing and Management (DISM) Tool...
Dism /online /cleanup-image /restorehealth
echo.

echo [8/11] Running Windows Disk Cleanup Utility...
cleanmgr /d C: /sagerun:1
echo.

echo [9/11] Clearing User Temporary Files...
rmdir /s /q "%temp%"
echo.

REM --- 3. DISK MAINTENANCE ---

echo [10/11] Running Disk Defragmenter (Defrag C:)...
defrag C: /U /V
echo.

echo [11/11] Running Disk Check (CHKDSK)...
echo CHKDSK requires a reboot to run. Running scan now.
chkdsk C: /f /r
echo.

echo ===========================================
echo   System Slowness Troubleshooting Complete.
echo ===========================================
pause