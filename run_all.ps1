$ErrorActionPreference = "Continue"
$Python = "C:\Users\kelland zhao\scoop\apps\python311\current\python.exe"
$Root = "C:\Users\kelland zhao\Projects\SAP_Google_AutoRun"
$LogFile = Join-Path $Root "run_all.log"

$Projects = @(
    "00 - Maintenance_Plan_Adherence_Sync",
    "01 - Maintenance_Effectiveness_Sync",
    "02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync",
    "03 - Safety_Stock_ZSE16",
    "04 - Total_Workorder",
    "05 - Stock_Turnover",
    "06 - IM_Equipment_Number",
    "07 - Inventory_MB52"
)

$Total = $Projects.Count
$Success = 0
$Failed = 0
$FailedList = @()

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "========================================"
Write-Host "  SAP 自动运行 - 共 $Total 个项目"
Write-Host "  开始时间: $Timestamp"
Write-Host "========================================"
Write-Host ""
"[$Timestamp] ========== 开始执行全部项目 ==========" | Out-File -Append $LogFile -Encoding UTF8

$Index = 0
foreach ($Project in $Projects) {
    $Index++
    $ProjectPath = Join-Path $Root $Project
    $MainPy = Join-Path $ProjectPath "main.py"
    
    if (-not (Test-Path $MainPy)) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 跳过 $Project : main.py 不存在"
        Write-Host $Msg -ForegroundColor Yellow
        $Msg | Out-File -Append $LogFile -Encoding UTF8
        $Failed++
        $FailedList += $Project
        continue
    }
    
    $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 执行中 $Project ..."
    Write-Host $Msg
    $Msg | Out-File -Append $LogFile -Encoding UTF8
    
    $Process = Start-Process -FilePath $Python -ArgumentList "`"$MainPy`"" -WorkingDirectory $ProjectPath -Wait -NoNewWindow -PassThru
    
    if ($Process.ExitCode -eq 0) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 完成 $Project"
        Write-Host $Msg -ForegroundColor Green
        $Success++
    } else {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 失败 $Project (退出码: $($Process.ExitCode))"
        Write-Host $Msg -ForegroundColor Red
        $Failed++
        $FailedList += $Project
    }
    $Msg | Out-File -Append $LogFile -Encoding UTF8
    Write-Host ""
}

$EndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "========================================"
Write-Host "  执行总结"
Write-Host "  总计: $Total | 成功: $Success | 失败: $Failed"
if ($FailedList.Count -gt 0) {
    Write-Host "  失败项目:"
    foreach ($f in $FailedList) {
        Write-Host "    - $f" -ForegroundColor Red
    }
}
Write-Host "  结束时间: $EndTime"
Write-Host "========================================"

$Msg = "[$EndTime] ========== 全部项目执行完毕 (成功: $Success, 失败: $Failed) =========="
$Msg | Out-File -Append $LogFile -Encoding UTF8
