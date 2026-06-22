$ErrorActionPreference = "Continue"
$Python = "C:\Users\kelland zhao\scoop\apps\python311\current\python.exe"
$Root = "C:\Users\kelland zhao\Desktop\072 - SAP自动程序源代码"
$LogFile = Join-Path $Root "run_all.log"

$Projects = @(
    "00 - Maintenance_Plan_Adherence_Sync",
    "01 - Maintenance_Effectiveness_Sync",
    "02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync",
    "03 - Safety_Stock_ZSE16",
    "04 - Total_Workorder",
    "05 - Stock_Turnover",
    "06 - IM_Equipment_Number",
    "07 - Inventory_MB52",
    "08 - WorkOrderAutoAcquisition"
)

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"[$Timestamp] ========== 开始执行全部项目 ==========" | Out-File -Append $LogFile -Encoding UTF8

foreach ($Project in $Projects) {
    $ProjectPath = Join-Path $Root $Project
    $MainPy = Join-Path $ProjectPath "main.py"
    
    if (-not (Test-Path $MainPy)) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ❌ 跳过 $Project : main.py 不存在"
        Write-Host $Msg
        $Msg | Out-File -Append $LogFile -Encoding UTF8
        continue
    }
    
    $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ▶ 开始执行 $Project"
    Write-Host $Msg
    $Msg | Out-File -Append $LogFile -Encoding UTF8
    
    $Process = Start-Process -FilePath $Python -ArgumentList "`"$MainPy`"" -WorkingDirectory $ProjectPath -Wait -NoNewWindow -PassThru
    
    if ($Process.ExitCode -eq 0) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ✅ 完成 $Project"
    } else {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ⚠️ $Project 退出码: $($Process.ExitCode)"
    }
    Write-Host $Msg
    $Msg | Out-File -Append $LogFile -Encoding UTF8
}

$Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ========== 全部项目执行完毕 =========="
Write-Host $Msg
$Msg | Out-File -Append $LogFile -Encoding UTF8
