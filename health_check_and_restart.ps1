# 机器人服务健康检查与端口映射修复脚本
# 适用于 Docker Desktop + WSL2 环境，解决容器 healthy 但主机端口不可访问的问题

$services = @(
    @{ Name = "hardware-summary"; Port = 8001; ComposeDir = "D:\Services\robot-services\hardware-summary\deploy\docker" },
    @{ Name = "order-split";       Port = 8002; ComposeDir = "D:\Services\robot-services\order-split\deploy\docker" },
    @{ Name = "dealer-sales";      Port = 8003; ComposeDir = "D:\Services\robot-services\may-sales\deploy\docker" },
    @{ Name = "csv-board";         Port = 8004; ComposeDir = "D:\Services\robot-services\csv-board\deploy\docker" },
    @{ Name = "pvc-classify";      Port = 8005; ComposeDir = "D:\Services\robot-services\pvc-classify\deploy\docker" },
    @{ Name = "workshop-order";    Port = 8006; ComposeDir = "D:\Services\robot-services\workshop-order\deploy\docker" }
)

$fixed = @()
$failed = @()

foreach ($svc in $services) {
    $port = $svc.Port
    $url = "http://localhost:$port/health"
    try {
        $resp = Invoke-WebRequest -Uri $url -TimeoutSec 5 -UseBasicParsing
        if ($resp.StatusCode -eq 200) {
            Write-Host "[OK] $($svc.Name) :$port 正常"
            continue
        }
    }
    catch {
        Write-Warning "[FAIL] $($svc.Name) :$port 无法访问，准备重启容器..."
    }

    try {
        Set-Location $svc.ComposeDir
        & docker compose -f docker-compose.yml down 2>&1 | Out-Null
        Start-Sleep -Seconds 2
        & docker compose -f docker-compose.yml up -d 2>&1 | Out-Null
        Start-Sleep -Seconds 5

        # 再次检查
        $resp2 = Invoke-WebRequest -Uri $url -TimeoutSec 5 -UseBasicParsing
        if ($resp2.StatusCode -eq 200) {
            Write-Host "[FIXED] $($svc.Name) :$port 已恢复" -ForegroundColor Green
            $fixed += $svc.Name
        }
        else {
            throw "重启后仍无法访问"
        }
    }
    catch {
        Write-Error "[ERROR] $($svc.Name) :$port 修复失败: $_"
        $failed += $svc.Name
    }
}

if ($fixed.Count -gt 0) {
    Write-Host "`n已修复服务: $($fixed -join ', ')" -ForegroundColor Green
}
if ($failed.Count -gt 0) {
    Write-Host "`n修复失败服务: $($failed -join ', ')" -ForegroundColor Red
}
