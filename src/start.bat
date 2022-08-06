@echo off

:: this script is used to start the main program so
:: that the user does not need to run it through the CLI

cd "%~dp0"
if exist "AutoGpuAffinity.py" (
    "AutoGpuAffinity.py"
) else (
    if exist "AutoGpuAffinity.exe" (
        "AutoGpuAffinity.exe"
    )
)

pause
exit /b 0
