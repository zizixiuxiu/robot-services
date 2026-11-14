# 2026-11-11 17:30 禁用的本地任务恢复脚本（生产已迁服务器，本地不再需要）
# 用法: 右键"使用 PowerShell 运行"，或 pwsh -File 本文件

Enable-ScheduledTask -TaskName 'Dealer-Sales-HTTP-Service' -ErrorAction SilentlyContinue | Out-Null; Write-Host '已恢复: Dealer-Sales-HTTP-Service'
Enable-ScheduledTask -TaskName 'PVC Trigger Service' -ErrorAction SilentlyContinue | Out-Null; Write-Host '已恢复: PVC Trigger Service'
Enable-ScheduledTask -TaskName 'PVC-Classify-HTTP-Service' -ErrorAction SilentlyContinue | Out-Null; Write-Host '已恢复: PVC-Classify-HTTP-Service'
Enable-ScheduledTask -TaskName 'PVC-鑷姩鍒嗙被-HTTP鏈嶅姟' -ErrorAction SilentlyContinue | Out-Null; Write-Host '已恢复: PVC-鑷姩鍒嗙被-HTTP鏈嶅姟'
Enable-ScheduledTask -TaskName 'RobotServicesStartup' -ErrorAction SilentlyContinue | Out-Null; Write-Host '已恢复: RobotServicesStartup'
Enable-ScheduledTask -TaskName 'workshop-ishare-sync' -ErrorAction SilentlyContinue | Out-Null; Write-Host '已恢复: workshop-ishare-sync'
