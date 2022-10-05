@echo off
setlocal EnableDelayedExpansion

:: this script is used to start the main program so
:: that the user does not need to run it through the CLI

pushd "%~dp0"
set "program_name=AutoGpuAffinity"
if exist ".\!program_name!.py" (
    ".\!program_name!.py"
) else (
    if exist ".\!program_name!.exe" (
        ".\!program_name!.exe"
    ) else (
        echo error: !program_name! not found
    )
)

echo info: press any key to continue
pause > nul 2>&1
exit /b 0
