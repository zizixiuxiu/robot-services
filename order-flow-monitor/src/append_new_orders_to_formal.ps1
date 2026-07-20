$ErrorActionPreference = "Stop"

$BaseDir = if ($env:ORDER_FLOW_BASE_DIR) { $env:ORDER_FLOW_BASE_DIR } else { "D:\Services\robot-services\order-flow-monitor" }
$WorkDir = if ($env:ORDER_FLOW_WORKDIR) { $env:ORDER_FLOW_WORKDIR } else { Join-Path $BaseDir "src" }
$Python = "C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$Node = "C:\Program Files\nodejs\node.exe"
$ExportScript = Join-Path $WorkDir "export_kdocs_online_workbook.js"
$ExportRoot = Join-Path (Join-Path $BaseDir "data") "order_flow_exports"
$OnlineExportPath = Join-Path $ExportRoot "latest_online_export.xlsx"
$AlertLog = Join-Path (Join-Path $BaseDir "logs") "order_flow_alerts.log"

$WorkbookName = "$([char]0x8BA2)$([char]0x5355)$([char]0x6D41)$([char]0x8F6C)$([char]0x8868)$([char]0xFF08)6-12$([char]0x6708)$([char]0xFF09)$([char]0xFF08)$([char]0x6B63)$([char]0x5F0F)$([char]0x7248)$([char]0xFF09).xlsx"
$OnlineFolder = "WPS$([char]0x8BA2)$([char]0x5355)$([char]0x6D41)$([char]0x8F6C)$([char]0x8868)$([char]0x5728)$([char]0x7EBF)$([char]0x8868)$([char]0x683C)"
$WpsCloudFolder = "WPS$([char]0x4E91)$([char]0x76D8)"
$LuxuryFolder = "$([char]0x5962)$([char]0x4F88)"
$OrderFlowFolder = "$([char]0x8BA2)$([char]0x5355)$([char]0x6D41)$([char]0x4F20)$([char]0x8868)"
$MigratedOrderFlowFolder = "$([char]0x8BA2)$([char]0x5355)$([char]0x6D41)$([char]0x8F6C)$([char]0x8868)$([char]0xFF08)6-12$([char]0x6708)$([char]0xFF09)"
$BackupRoot = Join-Path (Join-Path $BaseDir "data") "order_flow_backups"
$StableSeconds = 30
$StableTimeoutSeconds = 120

$Targets = @(
    (Join-Path (Join-Path "C:\Users\Administrator\WPSDrive\474560620" (Join-Path $WpsCloudFolder $MigratedOrderFlowFolder)) $WorkbookName)
)

if (-not (Test-Path -LiteralPath $BackupRoot)) {
    New-Item -ItemType Directory -Path $BackupRoot | Out-Null
}
if (-not (Test-Path -LiteralPath $ExportRoot)) {
    New-Item -ItemType Directory -Path $ExportRoot | Out-Null
}

$UpdatedTargets = 0

function Wait-WorkbookStable {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$StableSeconds = 30,
        [int]$TimeoutSeconds = 120
    )

    $Deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $Item = Get-Item -LiteralPath $Path
    $LastLength = $Item.Length
    $LastWriteUtc = $Item.LastWriteTimeUtc
    $StableSince = Get-Date

    while ((Get-Date) -lt $Deadline) {
        Start-Sleep -Seconds 5
        $Item = Get-Item -LiteralPath $Path
        if ($Item.Length -ne $LastLength -or $Item.LastWriteTimeUtc -ne $LastWriteUtc) {
            $LastLength = $Item.Length
            $LastWriteUtc = $Item.LastWriteTimeUtc
            $StableSince = Get-Date
            Write-Host "Workbook is still syncing/changing, waiting: $Path"
            continue
        }

        if (((Get-Date) - $StableSince).TotalSeconds -ge $StableSeconds) {
            return $true
        }
    }

    return $false
}

function Send-RefreshAlert {
    param(
        [Parameter(Mandatory=$true)][string]$Reason,
        [string]$Detail = ""
    )

    $TimeText = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Message = "Order flow auto refresh failed: " + $Reason + ". Please rerun KDocs login/export in the Playwright browser, then wait for the next refresh."
    if ($Detail) {
        $Message = $Message + " Detail: " + $Detail
    }

    Add-Content -LiteralPath $AlertLog -Encoding UTF8 -Value "$TimeText $Message"

    $Desktop = [Environment]::GetFolderPath("Desktop")
    if ($Desktop) {
        $AlertFile = Join-Path $Desktop "order_flow_auto_refresh_alert.txt"
        Set-Content -LiteralPath $AlertFile -Encoding UTF8 -Value $Message
    }

    try {
        & "$env:SystemRoot\System32\msg.exe" * /time:120 $Message | Out-Null
    } catch {
        Write-Warning "refresh_status=alert_msg_failed message=$($_.Exception.Message)"
    }

    Write-Warning "refresh_status=alert_sent reason=$Reason"
}

foreach ($Target in $Targets) {
    if (-not (Test-Path -LiteralPath $Target)) {
        Write-Warning "refresh_status=skipped reason=target_workbook_not_found target=$Target"
        continue
    }

    if (-not (Wait-WorkbookStable -Path $Target -StableSeconds $StableSeconds -TimeoutSeconds $StableTimeoutSeconds)) {
        Write-Warning "refresh_status=skipped reason=wps_workbook_not_stable stable_seconds=$StableSeconds timeout_seconds=$StableTimeoutSeconds target=$Target"
        continue
    }

    $TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $SafeFormalName = [IO.Path]::GetFileNameWithoutExtension($Target)
    $FormalBackupTarget = Join-Path $BackupRoot "$SafeFormalName`_online_before_sql_sync_$TimeStamp.xlsx"
    if (Test-Path -LiteralPath $OnlineExportPath) {
        Remove-Item -LiteralPath $OnlineExportPath -Force
    }

    $ProfilePath = Join-Path $WorkDir ".kdocs-playwright-profile"
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object {
        $_.CommandLine -like "*$ProfilePath*"
    } | ForEach-Object {
        Write-Host "killing_existing_edge pid=$($_.ProcessId) profile=$ProfilePath"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2

    $PreviousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $ExportOutput = & $Node $ExportScript --output $OnlineExportPath 2>&1
        $ExportExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $PreviousErrorActionPreference
    }
    foreach ($Line in $ExportOutput) {
        Write-Host $Line
    }
    if ($ExportExitCode -ne 0) {
        $ExportDetail = ($ExportOutput -join " ")
        if ($ExportDetail -match "kdocs_login_required") {
            Send-RefreshAlert -Reason "kdocs_login_required" -Detail "WPS/KDocs login expired"
        } else {
            Send-RefreshAlert -Reason "kdocs_online_export_failed" -Detail $ExportDetail
        }
        throw "refresh_status=failed reason=kdocs_online_export_failed exit_code=$ExportExitCode"
    }
    if (-not (Test-Path -LiteralPath $OnlineExportPath)) {
        throw "refresh_status=failed reason=kdocs_online_export_missing path=$OnlineExportPath"
    }
    if (-not (Wait-WorkbookStable -Path $OnlineExportPath -StableSeconds 5 -TimeoutSeconds 60)) {
        throw "refresh_status=failed reason=kdocs_online_export_not_stable path=$OnlineExportPath"
    }

    Copy-Item -LiteralPath $OnlineExportPath -Destination $FormalBackupTarget -Force
    Write-Host "refresh_status=formal_online_backup_created path=$FormalBackupTarget"

    $SafeTargetName = [IO.Path]::GetFileNameWithoutExtension($Target)
    $BackupTarget = Join-Path $BackupRoot "$SafeTargetName`_hidden_data_before_sql_sync_$TimeStamp.xlsx"
    $CopyRetry = 0
    $CopyMaxRetry = 5
    while ($CopyRetry -lt $CopyMaxRetry) {
        try {
            Copy-Item -LiteralPath $Target -Destination $BackupTarget -Force -ErrorAction Stop
            break
        } catch {
            $CopyRetry++
            if ($CopyRetry -ge $CopyMaxRetry) {
                throw "refresh_status=failed reason=local_backup_copy_failed message='$($_.Exception.Message)'"
            }
            Write-Warning "local_backup_copy_attempt=$CopyRetry retrying in 3s..."
            Start-Sleep -Seconds 3
        }
    }
    Write-Host "refresh_status=backup_created path=$BackupTarget"

    $env:TARGET_XLSX = $OnlineExportPath
    & $Python (Join-Path $WorkDir "append_new_sql_data_to_workbook.py")
    if ($LASTEXITCODE -ne 0) {
        throw "append_new_sql_data_to_workbook.py failed with exit code $LASTEXITCODE for $OnlineExportPath"
    }
    Copy-Item -LiteralPath $OnlineExportPath -Destination $Target -Force
    Write-Host "refresh_status=local_formal_replaced_from_online_export target=$Target source=$OnlineExportPath"
    $UpdatedTargets += 1
}

if ($UpdatedTargets -eq 0) {
    throw "refresh_status=failed reason=no_target_workbook_refreshed message='No target workbook was refreshed. This refresh will be retried by the monitor.'"
}
