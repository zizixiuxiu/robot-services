# Start and guard dealer-sales HTTP service on port 8003
$pythonExe = 'C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe'
$scriptPath = 'D:\Services\robot-services\may-sales\src\dealer_sales_http.py'
$port = 8003
$logPath = 'D:\Services\robot-services\may-sales\logs\dealer_sales_http.log'

function Stop-OldService {
    $oldProcs = Get-WmiObject Win32_Process -Filter "Name='python.exe'" | Where-Object { $_.CommandLine -like "*dealer_sales_http.py*" }
    foreach ($old in $oldProcs) {
        Stop-Process -Id $old.ProcessId -Force -ErrorAction SilentlyContinue
        Write-Host "[INFO] stopped old PID=$($old.ProcessId)"
    }
}

function Test-ServicePort {
    $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
    return ($null -ne $conn)
}

function Start-ServiceOnce {
    Stop-OldService
    Start-Sleep -Seconds 2
    # Redirect stdout/stderr to log file so the process can start without a console
    $argList = '"' + $pythonExe + '" "' + $scriptPath + '" --port ' + $port + ' >> "' + $logPath + '" 2>&1'
    Write-Host "[INFO] start command: $argList"
    $proc = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $argList }
    if ($proc.ReturnValue -eq 0) {
        Write-Host "[INFO] spawn ok PID=$($proc.ProcessId), waiting for port..."
        # Wait up to 15 seconds for the real python process to bind the port
        for ($i = 0; $i -lt 15; $i++) {
            Start-Sleep -Seconds 1
            if (Test-ServicePort) {
                Write-Host "[OK] port $port listening"
                return $true
            }
        }
        Write-Host "[WARN] port $port still not listening after 15s"
        return $false
    } else {
        Write-Host "[FAIL] spawn failed ReturnValue=$($proc.ReturnValue)"
        return $false
    }
}

# First start
if (-not (Start-ServiceOnce)) {
    Write-Host "[FAIL] initial start failed"
}

# Guard loop
while ($true) {
    if (-not (Test-ServicePort)) {
        Write-Host "[WARN] port $port not listening, restarting..."
        Start-ServiceOnce | Out-Null
    }
    Start-Sleep -Seconds 5
}
