# 启动当前在用的 Docker 服务（网关放最后）
$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot

$ComposeDirs = @(
    "hardware-summary\deploy\docker",
    "order-split\deploy\docker",
    "may-sales\deploy\docker",
    "csv-board\deploy\docker",
    "pvc-classify\deploy\docker",
    "workshop-order\deploy\docker",
    "quote-maker\deploy\docker",
    "dealer-report\deploy\docker",
    "attendance-summary\deploy\docker",
    "door-skin-converter",
    "order-summary\deploy\docker",
    "pms-door-split",
    "simple-ims\deploy\docker",
    "feishu-ws-gateway\deploy\docker"
)

$GatewayEnv = Join-Path $Root "feishu-ws-gateway\.env"
if (-not (Test-Path -LiteralPath $GatewayEnv)) {
    throw "先复制 feishu-ws-gateway\.env.example 为 .env 并填写 FEISHU_APP_SECRET"
}

foreach ($Rel in $ComposeDirs) {
    $Dir = Join-Path $Root $Rel
    Write-Host "=== starting $Rel ==="
    Push-Location $Dir
    try {
        docker compose up -d --build
    } finally {
        Pop-Location
    }
}

Write-Host "All services started. Check http://localhost:8001/health ... 8013/health"
