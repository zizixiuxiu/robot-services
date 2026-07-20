# switch.ps1 - 新旧版本一键切换（HTTP 层 _http.py）
# 用法:
#   powershell.exe -File switch.ps1 -Service all -Target old
#   powershell.exe -File switch.ps1 -Service csv-board -Target new
param(
    [Parameter(Mandatory = $true)]
    [string]$Service,

    [Parameter(Mandatory = $true)]
    [ValidateSet('old', 'new')]
    [string]$Target
)

$ErrorActionPreference = 'Stop'

# 服务清单: 服务名 -> 容器名 / 端口 / src 内 http 文件名 / version-switch 内文件名
$Services = [ordered]@{
    'hardware-summary' = @{ Container = 'hardware-summary-8001'; Port = 8001; SrcRel = 'hardware-summary\src\hardware_summary_http.py'; File = 'hardware_summary_http.py' }
    'order-split'      = @{ Container = 'order-split-8002';      Port = 8002; SrcRel = 'order-split\src\order_split_http.py';         File = 'order_split_http.py' }
    'may-sales'        = @{ Container = 'dealer-sales-8003';     Port = 8003; SrcRel = 'may-sales\src\dealer_sales_http.py';          File = 'dealer_sales_http.py' }
    'csv-board'        = @{ Container = 'csv-board-8004';        Port = 8004; SrcRel = 'csv-board\src\csv_board_http.py';             File = 'csv_board_http.py' }
    'pvc-classify'     = @{ Container = 'pvc-classify-8005';     Port = 8005; SrcRel = 'pvc-classify\src\pvc_classify_http.py';       File = 'pvc_classify_http.py' }
    'workshop-order'   = @{ Container = 'workshop-order-8006';   Port = 8006; SrcRel = 'workshop-order\src\workshop_order_http.py';   File = 'workshop_order_http.py' }
    'quote-maker'      = @{ Container = 'quote-maker-8007';      Port = 8007; SrcRel = 'quote-maker\src\quote_maker_http.py';         File = 'quote_maker_http.py' }
    'dealer-report'    = @{ Container = 'dealer-report-8008';    Port = 8008; SrcRel = 'dealer-report\src\dealer_report_http.py';     File = 'dealer_report_http.py' }
}

# 参数校验
if ($Service -ne 'all' -and -not $Services.Contains($Service)) {
    Write-Error "未知服务名 '$Service'。可选: all, $($Services.Keys -join ', ')"
    exit 1
}

$Root      = Split-Path -Parent $PSScriptRoot          # D:\Services\robot-services
$SourceDir = Join-Path $PSScriptRoot $Target           # version-switch\old 或 new
$BackupDir = Join-Path $PSScriptRoot 'backup'
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

if (-not (Test-Path $SourceDir)) { Write-Error "源目录不存在: $SourceDir"; exit 1 }
if (-not (Test-Path $BackupDir)) { New-Item -ItemType Directory -Path $BackupDir | Out-Null }

$targets = if ($Service -eq 'all') { $Services.Keys } else { @($Service) }

Write-Host "=== 切换目标: $Target | 服务: $($targets -join ', ') ===" -ForegroundColor Cyan

# 1) 备份当前文件 + 复制目标版本
foreach ($svc in $targets) {
    $info   = $Services[$svc]
    $src    = Join-Path $Root $info.SrcRel
    $source = Join-Path $SourceDir ($info.File + '.' + $Target)
    if (-not (Test-Path $source)) { Write-Error "缺少版本文件: $source"; exit 1 }
    if (Test-Path $src) {
        $backup = Join-Path $BackupDir ($info.File + '.' + $Timestamp)
        Copy-Item $src $backup -Force
        Write-Host "[$svc] 已备份当前文件 -> $backup"
    }
    Copy-Item $source $src -Force
    Write-Host "[$svc] 已复制 $Target 版本 -> $src"
}

# 2) 重启容器 + 轮询 health
$results = @()
foreach ($svc in $targets) {
    $info = $Services[$svc]
    Write-Host "[$svc] docker restart $($info.Container) ..."
    docker restart $info.Container | Out-Null

    $healthy = $false
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline) {
        try {
            $resp = Invoke-WebRequest -Uri "http://localhost:$($info.Port)/health" -UseBasicParsing -TimeoutSec 3
            if ($resp.StatusCode -eq 200) { $healthy = $true; break }
        } catch { Start-Sleep -Seconds 2 }
    }
    $status = if ($healthy) { 'OK' } else { 'TIMEOUT(60s)' }
    $results += [pscustomobject]@{ Service = $svc; Container = $info.Container; Port = $info.Port; Target = $Target; Health = $status }
}

Write-Host ""
Write-Host "=== 切换结果 ===" -ForegroundColor Cyan
$results | Format-Table -AutoSize

if ($results | Where-Object { $_.Health -ne 'OK' }) {
    Write-Host "警告: 有服务 health 检查未通过!" -ForegroundColor Red
    exit 2
}
Write-Host "全部服务切换完成, health 正常。" -ForegroundColor Green
