$ErrorActionPreference = "Continue"
$Python = "C:\Users\kelland zhao\scoop\apps\python311\current\python.exe"
$Root = "C:\Users\kelland zhao\Projects\SAP_Google_AutoRun"
$LogFile = Join-Path $Root "run_all.log"

# 当前代码版本 —— 写进日志，便于事后核对本次跑的是哪一版
$GitRev = "unknown"
if (Get-Command git -ErrorAction SilentlyContinue) {
    $GitRev = (git -C "$Root" rev-parse --short HEAD 2>$null | Select-Object -First 1)
    if (-not $GitRev) { $GitRev = "unknown" }
}

$Projects = @(
    "00 - Maintenance_Plan_Adherence_Sync",
    "01 - Maintenance_Effectiveness_Sync",
    "02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync",
    "03 - Safety_Stock_ZSE16",
    "04 - Total_Workorder",
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
Write-Host "  代码版本: $GitRev"
Write-Host "  开始: $Timestamp"
Write-Host "========================================"
Write-Host ""
"[$Timestamp] ========== 开始执行 (代码版本: $GitRev) ==========" | Out-File -Append $LogFile -Encoding UTF8

# ── 启动 SAP（只启动一次）────────────────────────────────────────────────
# 各脚本检测到 SAP 已在运行就不再启停，改成在自己的新会话里干活。
# 这一步省掉原先"每个脚本重启一次 SAP"的 7 次开销。
Write-Host "  正在启动 SAP（各脚本将复用，不再逐个重启）..."
& $Python (Join-Path $Root "sap_common.py") start
if ($LASTEXITCODE -ne 0) {
    Write-Host "  ❌ SAP 启动失败，本次运行终止（否则 7 个脚本会各等 60 秒再全部失败）。" -ForegroundColor Red
    "[$Timestamp] SAP 启动失败，运行终止" | Out-File -Append $LogFile -Encoding UTF8
    exit 1
}

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
    
    $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 执行 $Project ..."
    Write-Host $Msg
    $Msg | Out-File -Append $LogFile -Encoding UTF8
    
    # 必须保留 -Wait：PowerShell 里 Start-Process 不带 -Wait 时拿不到 ExitCode，
    # 曾因此导致 7 个脚本全部被误报为"失败 (退出码: )"。
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

# ── 关闭 SAP（本编排器启动的那一个）──────────────────────────────────────
Write-Host "  正在关闭 SAP..."
& $Python (Join-Path $Root "sap_common.py") close

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
Write-Host "  结束: $EndTime"
Write-Host "========================================"

"[$EndTime] ========== 全部完成 (成功: $Success, 失败: $Failed) ==========" | Out-File -Append $LogFile -Encoding UTF8

# 注意：此处原先会执行 `taskkill /f /im excel.exe`，强制关闭机器上所有 Excel。
# 这会杀掉同事或自己正在编辑、尚未保存的工作簿，已移除。
# 各脚本内部已有 wb.Close(SaveChanges=False) + excel.Quit() 负责收尾。

Write-Host ""
