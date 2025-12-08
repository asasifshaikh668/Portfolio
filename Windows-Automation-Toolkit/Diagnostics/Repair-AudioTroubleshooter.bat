@echo off

echo Diagnosing Sound and Microphone Issues...

REM Run the built-in audio troubleshooter
msdt.exe /id AudioPlaybackDiagnostic

REM Run the built-in microphone troubleshooter
msdt.exe /id AudioRecordingDiagnostic

REM Restart the Windows Audio service
net stop Audiosrv
net start Audiosrv

REM Reset the audio device
for /F "tokens=*" %%A in ('powershell -command "(Get-PnpDevice | Where-Object { $_.Status -eq 'Error' -and $_.ClassGuid -eq '4d36e96c-e325-11ce-bfc1-08002be10318' }).InstanceId"') do (
    powershell -command "Disable-PnpDevice -InstanceId '%%A' -Confirm:$false"
    powershell -command "Enable-PnpDevice -InstanceId '%%A' -Confirm:$false"
)

REM Re-register core audio components
regsvr32 /s audiosrv.dll
regsvr32 /s audiocore.dll

echo Sound and Microphone Troubleshooting Complete.
pause
