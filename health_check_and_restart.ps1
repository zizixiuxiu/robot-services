# Health check and targeted repair for robot services.
# New entry: Docker services live under D:\Services\robot-services.
# Exception: OrderFlowMonitor intentionally uses the legacy D:\1 entry.
$ErrorActionPreference = "SilentlyContinue"

$services = @(
    @{ Name = "hardware-summary"; Port = 8001; ComposeDir = "D:\Services\robot-services\hardware-summary\deploy\docker" },
    @{ Name = "order-split";       Port = 8002; ComposeDir = "D:\Services\robot-services\order-split\deploy\docker" },
    @{ Name = "dealer-sales";      Port = 8003; ComposeDir = "D:\Services\robot-services\may-sales\deploy\docker" },
    @{ Name = "csv-board";         Port = 8004; ComposeDir = "D:\Services\robot-services\csv-board\deploy\docker" },
    @{ Name = "pvc-classify";      Port = 8005; ComposeDir = "D:\Services\robot-services\pvc-classify\deploy\docker" },
    @{ Name = "workshop-order";    Port = 8006; ComposeDir = "D:\Services\robot-services\workshop-order\deploy\docker" },
    @{ Name = "quote-maker";       Port = 8007; ComposeDir = "D:\Services\robot-services\quote-maker\deploy\docker" },
    @{ Name = "simple-ims";        Port = 8090; ComposeDir = "D:\Services\robot-services\simple-ims\deploy\docker" }
)

function Restart-DockerCompose($composeDir, $name) {
    Push-Location $composeDir
    try {
        docker compose -f docker-compose.yml up -d 2>&1 | Out-Null
        Write-Host "[FIXED] $name started" -ForegroundColor Green
    } catch {
        Write-Error "[ERROR] $name repair failed: $($_.Exception.Message)"
    } finally {
        Pop-Location
    }
}

foreach ($svc in $services) {
    $url = "http://127.0.0.1:$($svc.Port)/health"
    try {
        $resp = Invoke-WebRequest -Uri $url -TimeoutSec 5 -UseBasicParsing
        if ($resp.StatusCode -eq 200) {
            Write-Host "[OK] $($svc.Name) :$($svc.Port) healthy"
            continue
        }
    } catch {
        Write-Warning "[FAIL] $($svc.Name) :$($svc.Port) unavailable; restarting container..."
    }

    Restart-DockerCompose $svc.ComposeDir $svc.Name
}

$gatewayRunning = $false
$gatewayState = docker inspect --format "{{.State.Running}}" feishu-ws-gateway 2>$null
if ($gatewayState -eq "true") {
    $gatewayRunning = $true
}

if ($gatewayRunning) {
    Write-Host "[OK] feishu-ws-gateway container running"
} else {
    Write-Warning "[FAIL] feishu-ws-gateway container is not running; restarting..."
    Restart-DockerCompose "D:\Services\robot-services\feishu-ws-gateway\deploy\docker" "feishu-ws-gateway"
}

$orderFlow = Get-WmiObject Win32_Process -Filter "Name='python.exe'" | Where-Object {
    $_.CommandLine -match "C:\\Users\\Administrator\\Documents\\Codex\\2026-05-28\\sqlserver\\monitor_refresh_order_flow.py"
} | Select-Object -First 1

if ($orderFlow) {
    Write-Host "[OK] OrderFlowMonitor legacy entry running PID=$($orderFlow.ProcessId)"
} else {
    Write-Warning "[FAIL] OrderFlowMonitor is not running; starting legacy entry..."
    & "C:\Windows\System32\wscript.exe" "D:\1\start_order_flow_monitor.vbs"
}
