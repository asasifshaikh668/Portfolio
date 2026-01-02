function Test-RemoteConnectivity {
    $Target = Read-Host "Enter Hostname or IP Address"
    $Ports = @(80, 443, 3389)

    Write-Host "`nTesting ICMP Ping for $Target..." -ForegroundColor Cyan
    Test-Connection -ComputerName $Target -Count 1 -Quiet

    foreach ($Port in $Ports) {
        Write-Host "Testing TCP Port $Port..." -ForegroundColor Cyan
        # Test-NetConnection verifies if specific ports are open [cite: 1, 5]
        Test-NetConnection -ComputerName $Target -Port $Port | 
        Select-Object ComputerName, RemotePort, TcpTestSucceeded | 
        Format-Table -AutoSize
    }
}

Test-RemoteConnectivity