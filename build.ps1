function main() {
    # pack executable
    pyinstaller ".\AutoGpuAffinity\main.py" --onefile --name AutoGpuAffinity

    if (Test-Path ".\build\") {
        Remove-Item -Path ".\build\" -Recurse
    }

    # create folder structure
    New-Item -ItemType Directory -Path ".\build\AutoGpuAffinity\"

    # create final package
    Move-Item ".\dist\AutoGpuAffinity.exe" ".\build\AutoGpuAffinity\"
    Move-Item ".\AutoGpuAffinity\bin\" ".\build\AutoGpuAffinity\"
    Move-Item ".\AutoGpuAffinity\config.ini" ".\build\AutoGpuAffinity\"

    return 0
}

$_exitCode = main
Write-Host # new line
exit $_exitCode
