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

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"[$Timestamp] ========== START ALL PROJECTS ==========" | Out-File -Append $LogFile -Encoding UTF8

foreach ($Project in $Projects) {
    $ProjectPath = Join-Path $Root $Project
    $MainPy = Join-Path $ProjectPath "main.py"
    
    if (-not (Test-Path $MainPy)) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SKIP $Project : main.py not found"
        Write-Host $Msg
        $Msg | Out-File -Append $LogFile -Encoding UTF8
        continue
    }
    
    $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] RUN $Project"
    Write-Host $Msg
    $Msg | Out-File -Append $LogFile -Encoding UTF8
    
    $Process = Start-Process -FilePath $Python -ArgumentList "`"$MainPy`"" -WorkingDirectory $ProjectPath -Wait -NoNewWindow -PassThru
    
    if ($Process.ExitCode -eq 0) {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] DONE $Project"
    } else {
        $Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] FAIL $Project (exit code: $($Process.ExitCode))"
    }
    Write-Host $Msg
    $Msg | Out-File -Append $LogFile -Encoding UTF8
}

$Msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ========== ALL PROJECTS COMPLETED =========="
Write-Host $Msg
$Msg | Out-File -Append $LogFile -Encoding UTF8

Write-Host ""
Write-Host "All done."
