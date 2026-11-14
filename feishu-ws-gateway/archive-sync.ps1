# gateway-archive.ps1 v3 —— 对象存储版（效仿企业「存储与处理分离」）
# 流程（每 1 分钟由计划任务 GatewayArchiveSync 调用）：
#   1) 同步：archive\<群名>\<日期>\ 新文件（≥60s 龄，防半写）
#      → 服务器 archive-s3/staging/<群名>/<日期>/（幂等：mtime 标记增量，失败留待下轮）
#   2) 服务器侧（cron 每 5 分钟，独立）：staging → SeaweedFS S3 桶 gateway-archive
#      （sha256 + 尺寸校验后写元数据 JSONL、删 staging 副本；见 archive-s3/ingest.py）
#   3) 清理：本地副本经服务器 check.py 核实（S3 已入库 或 staging 有等尺寸副本）后，
#      保留 72 小时即自动删除；核实失败绝不删本地。
$ErrorActionPreference = "Continue"
$src = "D:\Services\robot-services\feishu-ws-gateway\logs\archive"
$dst = "/home/ops/robot-platform-tests/archive-s3/staging"
$marker = Join-Path $src ".last_sync"
$log = "D:\Services\robot-services\feishu-ws-gateway\logs\archive_sync.log"
$retentionHours = 72   # 本地中转副本保留时长（服务器核实通过后到期清理）

if (-not (Test-Path $src)) { exit 0 }
if (-not (Test-Path $marker)) {
    # 首跑：标记回拨到过去，确保已存在的存档文件全部被首轮同步
    New-Item -ItemType File -Path $marker -Force | Out-Null
    (Get-Item $marker).LastWriteTime = Get-Date "2000-01-01 00:00:00"
}
$lastMark = (Get-Item $marker).LastWriteTime
$cutoff = Get-Date
$minAge = $cutoff.AddSeconds(-60)

function Escape-RemotePath([string]$p) {
    return $p.Replace("'", "'\''")
}
# SSH 保活：连接死掉时 30 秒内失败退出（防止挂死实例阻塞每分钟任务）
$sshOpts = @("-p","2222","-o","BatchMode=yes","-o","ConnectTimeout=15","-o","ServerAliveInterval=10","-o","ServerAliveCountMax=3")
$scpOpts = @("-O","-P","2222","-o","BatchMode=yes","-o","ConnectTimeout=15","-o","ServerAliveInterval=10","-o","ServerAliveCountMax=3")

# ---------- 1) 同步到服务器 staging ----------
$files = Get-ChildItem -Path $src -Recurse -File |
    Where-Object { $_.Name -ne ".last_sync" -and $_.LastWriteTime -gt $lastMark -and $_.LastWriteTime -lt $minAge }

$ok = 0; $fail = 0
foreach ($f in $files) {
    $reldir = $f.DirectoryName.Substring($src.Length).TrimStart('\').Replace('\', '/')
    if ($reldir -eq "") { $reldir = "." }
    & ssh @sshOpts ops@192.168.6.245 "mkdir -p '$dst/$(Escape-RemotePath $reldir)'" 2>$null
    & scp @scpOpts "$($f.FullName)" "ops@192.168.6.245:$dst/$(Escape-RemotePath $reldir)/" 2>$null
    if ($LASTEXITCODE -eq 0) { $ok++ } else { $fail++; "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') SYNC_FAIL: $($f.FullName)" | Add-Content $log }
}
if ($files.Count -gt 0) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') synced=$ok fail=$fail" | Add-Content $log
}

# ---------- 2) 清理（服务器核实：S3 已入库 或 staging 等尺寸副本） ----------
$purged = 0
$stale = Get-ChildItem -Path $src -Recurse -File |
    Where-Object { $_.Name -ne ".last_sync" -and $_.LastWriteTime -lt (Get-Date).AddHours(-$retentionHours) }
foreach ($f in $stale) {
    $reldir = $f.DirectoryName.Substring($src.Length).TrimStart('\').Replace('\', '/')
    if ($reldir -eq "") { $reldir = "." }
    $out = & ssh @sshOpts ops@192.168.6.245 `
        "python3 /home/ops/robot-platform-tests/archive-s3/check.py '$(Escape-RemotePath $reldir)' '$(Escape-RemotePath $f.Name)' $($f.Length)" 2>$null
    if ($LASTEXITCODE -eq 0 -and "$out" -match 'OK') {
        Remove-Item -LiteralPath $f.FullName -Force
        $purged++
    } else {
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') PURGE_SKIP(服务器未核实): $($f.FullName)" | Add-Content $log
    }
}
# 清理空目录（只删 archive 下确实为空的目录）
Get-ChildItem -Path $src -Recurse -Directory | Sort-Object FullName -Descending |
    Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Force) } |
    Remove-Item -Force -ErrorAction SilentlyContinue
if ($purged -gt 0) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') purged_local=$purged (retention=${retentionHours}h)" | Add-Content $log
}

# 标记 = 本轮扫描开始时间；写于扫描之后的新文件下轮必被拾起
(Get-Item $marker).LastWriteTime = $cutoff
