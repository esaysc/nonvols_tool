# 设置关机时间为 17:10 (下午 5:10)
$Time = "9:36"
$TaskName = "DailyAutoShutdown"
$Action = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/s /f /t 60"
$Trigger = New-ScheduledTaskTrigger -Daily -At $Time
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# 检查任务是否已存在，如果存在则先删除
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# 注册新任务
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description "每天下午 5:10 自动关机"

Write-Host "成功创建任务：将在每天 $Time 自动关机（提前60秒提示）。"
