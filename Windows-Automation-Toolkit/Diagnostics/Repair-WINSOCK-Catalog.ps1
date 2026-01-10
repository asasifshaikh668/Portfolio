<#
.SYNOPSIS
    Resets the Winsock Catalog and TCP/IP stack to repair network connectivity.
.DESCRIPTION
    This script runs 'netsh winsock reset' and 'netsh int ip reset' to clear 
    network corruption caused by malware, driver issues, or configuration errors.
    It requires Administrator privileges and a system restart.
#>

# Ensure script is run as Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You must run this script as an Administrator!"
    Break
}

Write-Host "Resetting Winsock Catalog..." -ForegroundColor Cyan
netsh winsock reset

Write-Host "Resetting TCP/IP Stack..." -ForegroundColor Cyan
netsh int ip reset

Write-Host "`nOperation Complete. A restart is required." -ForegroundColor Yellow
$Restart = Read-Host "Do you want to restart now? (Y/N)"
if ($Restart -eq 'Y') {
    Restart-Computer -Force
}