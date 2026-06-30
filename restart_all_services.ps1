# Restart robot services from the new unified entry.
# OrderFlowMonitor is intentionally left on its legacy D:\1 entry.
$ErrorActionPreference = "SilentlyContinue"

$root = "D:\Services\robot-services"
$logDir = Join-Path $root "logs"
$logPath = Join-Path $logDir "restart_all_services.log"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Write-Log($msg) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp $msg" | Add-Content -Path $logPath
    Write-Host "[$timestamp] $msg"
}

Write-Log "[INFO] Stopping new-entry robot services..."

$legacyPatterns = @(
    "feishu_bot_ws.py",
    "uvicorn.*8090"
)

Get-WmiObject Win32_Process -Filter "Name='python.exe'" | Where-Object {
    $cmd = $_.CommandLine
    foreach ($pat in $legacyPatterns) {
        if ($cmd -match $pat) { return $true }
    }
    return $false
} | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force
    Write-Log "[INFO] stopped legacy PID=$($_.ProcessId) cmd=$($_.CommandLine.Substring(0, [Math]::Min(100, $_.CommandLine.Length)))"
}

$dockerServices = @(
    @("8001 hardware-summary", "D:\Services\robot-services\hardware-summary\deploy\docker\docker-compose.yml"),
    @("8002 order-split", "D:\Services\robot-services\order-split\deploy\docker\docker-compose.yml"),
    @("8003 may-sales", "D:\Services\robot-services\may-sales\deploy\docker\docker-compose.yml"),
    @("8004 csv-board", "D:\Services\robot-services\csv-board\deploy\docker\docker-compose.yml"),
    @("8005 pvc-classify", "D:\Services\robot-services\pvc-classify\deploy\docker\docker-compose.yml"),
    @("8006 workshop-order", "D:\Services\robot-services\workshop-order\deploy\docker\docker-compose.yml"),
    @("8007 quote-maker", "D:\Services\robot-services\quote-maker\deploy\docker\docker-compose.yml"),
    @("8090 simple-ims", "D:\Services\robot-services\simple-ims\deploy\docker\docker-compose.yml"),
    @("feishu-ws-gateway", "D:\Services\robot-services\feishu-ws-gateway\deploy\docker\docker-compose.yml")
)

foreach ($svc in $dockerServices) {
    $name = $svc[0]
    $file = $svc[1]
    Write-Log "[INFO] stopping Docker $name..."
    docker compose -f "$file" down 2>&1 | ForEach-Object { Write-Log "[INFO] docker: $_" }
}

Start-Sleep -Seconds 2

foreach ($svc in $dockerServices) {
    $name = $svc[0]
    $file = $svc[1]
    Write-Log "[INFO] starting Docker $name..."
    docker compose -f "$file" up -d 2>&1 | ForEach-Object { Write-Log "[INFO] docker: $_" }
}

Write-Log "[INFO] ensuring legacy OrderFlowMonitor entry is running..."
& "C:\Windows\System32\wscript.exe" "D:\1\start_order_flow_monitor.vbs"

$endpoints = @{
    8001 = "/health"
    8002 = "/health"
    8003 = "/health"
    8004 = "/health"
    8005 = "/health"
    8006 = "/health"
    8007 = "/health"
    8090 = "/health"
}

for ($attempt = 1; $attempt -le 60; $attempt++) {
    $allReady = $true
    foreach ($port in ($endpoints.Keys | Sort-Object)) {
        $uri = "http://127.0.0.1:${port}$($endpoints[$port])"
        try {
            $resp = Invoke-WebRequest -Uri $uri -UseBasicParsing -TimeoutSec 10
            Write-Log "[OK] port $port status=$($resp.StatusCode)"
        } catch {
            $allReady = $false
            if ($attempt -eq 60) {
                Write-Log "[FAIL] port $port not responding"
            }
        }
    }
    if ($allReady) { break }
    Start-Sleep -Seconds 5
}

Write-Log "[INFO] Restart complete."
