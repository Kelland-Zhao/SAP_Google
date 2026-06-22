# 创建 Windows Task Scheduler 定时任务
# 以管理员身份运行此脚本

$TaskName = "SAP_Auto_Run_All"
$ScriptPath = "C:\Users\kelland zhao\Desktop\072 - SAP自动程序源代码\run_all.ps1"
$PythonPath = "C:\Users\kelland zhao\scoop\apps\python311\current\python.exe"

# 删除旧任务（如果存在）
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

# 创建任务操作
$Action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -File `"$ScriptPath`""

# 创建触发器（每天 08:00）
$Trigger = New-ScheduledTaskTrigger -Daily -At "08:00"

# 任务配置：仅在用户登录时运行（SAP GUI 需要桌面会话）
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive -RunLevel Highest

# 设置：允许按需运行，超时后停止
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
    -MultipleInstances IgnoreNew

# 注册任务
Register-ScheduledTask -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal `
    -Settings $Settings `
    -Description "SAP 自动程序 00-08 逐个项目运行"

Write-Host "✅ 任务 '$TaskName' 已创建，每天 08:00 执行"
Write-Host "   脚本: $ScriptPath"
Write-Host ""
Write-Host "手动测试运行:"
Write-Host "   Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
Write-Host "修改执行时间: 打开 taskschd.msc → 找到 '$TaskName' → 属性 → 触发器"
