@echo off
setlocal

:MENU
cls
echo ===============================================
echo  GLOBAL TIME LOOKUP (Technical Support Tool)
echo ===============================================
echo [1] Australia (Sydney),[8] Dubai (GST)
echo [2] Singapore         ,[9] Egypt
echo [3] Hong Kong         ,[10] Germany
echo [4] China             ,[11] France
echo [5] Japan             ,[12] London (GMT/BST)
echo [6] North Korea       ,[13] New York (EST/EDT)
echo [7] South Korea       ,[14] India (IST)
echo -----------------------------------------------
echo [15] Exit Tool
echo -----------------------------------------------
set /p CHOICE="Enter your selection (1-15): "

if "%CHOICE%"=="1" goto AUST_SYD
if "%CHOICE%"=="2" goto SINGAPORE
if "%CHOICE%"=="3" goto HONGKONG
if "%CHOICE%"=="4" goto CHINA
if "%CHOICE%"=="5" goto JAPAN
if "%CHOICE%"=="6" goto KOREA_NORTH
if "%CHOICE%"=="7" goto KOREA_SOUTH
if "%CHOICE%"=="8" goto DUBAI
if "%CHOICE%"=="9" goto EGYPT
if "%CHOICE%"=="10" goto GERMANY
if "%CHOICE%"=="11" goto FRANCE
if "%CHOICE%"=="12" goto LONDON
if "%CHOICE%"=="13" goto NEWYORK
if "%CHOICE%"=="14" goto INDIA
if "%CHOICE%"=="15" goto END

echo Invalid choice. Please try again.
pause
goto MENU

:AUST_SYD
call :GETTIME "Australia (Sydney)" "AUS Eastern Standard Time"
goto MENU

:SINGAPORE
call :GETTIME "Singapore" "Singapore Standard Time"
goto MENU

:HONGKONG
call :GETTIME "Hong Kong" "Hong Kong Standard Time"
goto MENU

:CHINA
call :GETTIME "China" "China Standard Time"
goto MENU

:JAPAN
call :GETTIME "Japan" "Tokyo Standard Time"
goto MENU

:KOREA_NORTH
call :GETTIME "North Korea" "Korea Standard Time"
goto MENU

:KOREA_SOUTH
call :GETTIME "South Korea" "Korea Standard Time"
goto MENU

:INDIA
call :GETTIME "India (IST)" "India Standard Time"
goto MENU

:DUBAI
call :GETTIME "Dubai (GST)" "Arabian Standard Time"
goto MENU

:EGYPT
call :GETTIME "Egypt" "Egypt Standard Time"
goto MENU

:GERMANY
call :GETTIME "Germany" "W. Europe Standard Time"
goto MENU

:FRANCE
call :GETTIME "France" "W. Europe Standard Time"
goto MENU

:LONDON
call :GETTIME "London (GMT/BST)" "GMT Standard Time"
goto MENU

:NEWYORK
call :GETTIME "New York (EST/EDT)" "Eastern Standard Time"
goto MENU

:GETTIME
echo.
echo Current time in %~1:
powershell -command "[System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), '%~2') | Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"
pause
goto :EOF

:END
echo.
echo Exiting Global Time Lookup Tool.
endlocal