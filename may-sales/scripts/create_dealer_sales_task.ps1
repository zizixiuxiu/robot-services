$taskName = 'Dealer-Sales-HTTP-Service'
$Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-WindowStyle Hidden -ExecutionPolicy Bypass -File D:\Services\robot-services\may-sales\scripts\run_dealer_sales_service.ps1'
$TriggerLogon = New-ScheduledTaskTrigger -AtLogon
$TriggerOnDemand = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(-1)
$Principal = New-ScheduledTaskPrincipal -UserId 'Administrator' -LogonType Interactive -RunLevel Highest
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -Hidden
Register-ScheduledTask -TaskName $taskName -Action $Action -Trigger @($TriggerLogon, $TriggerOnDemand) -Principal $Principal -Settings $Settings -Force
Write-Host "Task updated: $taskName"
