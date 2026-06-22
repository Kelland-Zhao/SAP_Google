# Create Windows Task Scheduler scheduled task
# Run this script as Administrator

$TaskName = "SAP_Auto_Run_All"
$ScriptPath = "C:\Users\kelland zhao\Desktop\072 - SAP自动程序源代码\run_all.ps1"

# Remove old task if exists
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

# Action: run PowerShell script
$Action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -File `"$ScriptPath`""

# Trigger: daily at 09:05
$Trigger = New-ScheduledTaskTrigger -Daily -At "09:05"

# Principal: run only when user is logged on (SAP GUI requires desktop session)
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive -RunLevel Highest

# Settings
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
    -MultipleInstances IgnoreNew

# Register task
Register-ScheduledTask -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal `
    -Settings $Settings `
    -Description "SAP Auto Run All Projects 00-08"

Write-Host "Task '$TaskName' created, daily at 09:05"
Write-Host "Script: $ScriptPath"
Write-Host ""
Write-Host "Manual test: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Change time: taskschd.msc -> find '$TaskName' -> Properties -> Triggers"
