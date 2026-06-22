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
Write-Host "  SAP 鑷姩杩愯 - 鍏?$Total 涓」鐩?
Write-Host "  寮€濮? $Timestamp"
Write-Host "========================================"
Write-Host ""
"[$Timestamp] ========== 寮€濮嬫墽琛?==========" | Out-File -Append $LogFile -Encoding UTF8

$Index = 0
foreach ($Project in $Projects) {
    $Index++
    $ProjectPath = Join-Path $Root $Project
    $MainPy = Join-Path $ProjectPath "main.py"
    
    if (-not (Test-Path $MainPy)) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 璺宠繃 $Project : main.py 涓嶅瓨鍦?
        Write-Host $Msg -ForegroundColor Yellow
        $Msg | Out-File -Append $LogFile -Encoding UTF8
        $Failed++
        $FailedList += $Project
        continue
    }
    
    $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 鎵ц $Project ..."
    Write-Host $Msg
    $Msg | Out-File -Append $LogFile -Encoding UTF8
    
    $Process = Start-Process -FilePath $Python -ArgumentList "`"$MainPy`"" -WorkingDirectory $ProjectPath -Wait -NoNewWindow -PassThru
    
    if ($Process.ExitCode -eq 0) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 瀹屾垚 $Project"
        Write-Host $Msg -ForegroundColor Green
        $Success++
    } else {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Index/$Total] 澶辫触 $Project (閫€鍑虹爜: $($Process.ExitCode))"
        Write-Host $Msg -ForegroundColor Red
        $Failed++
        $FailedList += $Project
    }
    $Msg | Out-File -Append $LogFile -Encoding UTF8
    Write-Host ""
}

$EndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "========================================"
Write-Host "  鎵ц鎬荤粨"
Write-Host "  鎬昏: $Total | 鎴愬姛: $Success | 澶辫触: $Failed"
if ($FailedList.Count -gt 0) {
    Write-Host "  澶辫触椤圭洰:"
    foreach ($f in $FailedList) {
        Write-Host "    - $f" -ForegroundColor Red
    }
}
Write-Host "  缁撴潫: $EndTime"
Write-Host "========================================"

"[$EndTime] ========== 鍏ㄩ儴瀹屾垚 (鎴愬姛: $Success, 澶辫触: $Failed) ==========" | Out-File -Append $LogFile -Encoding UTF8
