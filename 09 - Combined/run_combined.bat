@echo off
chcp 65001 >nul
setlocal

set "PYTHON_EXE=C:\Users\kelland zhao\scoop\apps\python311\current\python.exe"

set "SCRIPT_DIR=%~dp0"
if not exist "%SCRIPT_DIR%" (
    echo [ERROR] Script directory not found: %SCRIPT_DIR%
    exit /b 1
)

pushd "%SCRIPT_DIR%" || (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    exit /b 1
)

set "SCRIPT_DIR=%CD%"
set "LOG_DIR=%SCRIPT_DIR%\logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

for /f %%i in ('powershell -NoProfile -Command "(Get-Date).ToString('yyyyMMdd_HHmmss')"') do set "STAMP=%%i"
set "LOG_FILE=%LOG_DIR%\combined_%STAMP%.log"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python executable not found: %PYTHON_EXE%
    popd
    exit /b 1
)

if not exist "%SCRIPT_DIR%\CombinedMain.py" (
    echo [ERROR] CombinedMain.py not found under %SCRIPT_DIR%
    popd
    exit /b 1
)

"%PYTHON_EXE%" "%SCRIPT_DIR%\CombinedMain.py" 

if errorlevel 1 (
    echo [ERROR] Python script exited with error. See log: %LOG_FILE%
) else (
    echo [INFO] Combined workflow completed. Log: %LOG_FILE%
)

echo.
echo 运行结束，请按任意键关闭窗口 / Run finished, press any key to close.
pause >nul

popd
endlocal