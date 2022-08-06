@echo off
setlocal EnableDelayedExpansion

:: this script is used to start the main program so
:: that the user does not need to run it through the CLI

pushd "%~dp0"
set "program_name=AutoGpuAffinity"
if exist "!program_name!.py" (
    "!program_name!.py"
) else (
    if exist "!program_name!.exe" (
        "!program_name!.exe"
    ) else (
        echo error: !program_name! not found
    )
)

pause
exit /b 0
