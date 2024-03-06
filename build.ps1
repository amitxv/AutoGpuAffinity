function main() {
    if (Test-Path ".\build\") {
        Remove-Item -Path ".\build\" -Recurse
    }

    # create folder structure
    New-Item -ItemType Directory -Path ".\build\AutoGpuAffinity\"

    # pack executable
    New-Item -ItemType Directory -Path ".\build\pyinstaller\"
    Push-Location ".\build\pyinstaller\"
    pyinstaller "..\..\AutoGpuAffinity\main.py" --onefile --name AutoGpuAffinity
    Pop-Location

    # create final package
    Copy-Item ".\build\pyinstaller\dist\AutoGpuAffinity.exe" ".\build\AutoGpuAffinity\"
    Copy-Item ".\AutoGpuAffinity\bin\" ".\build\AutoGpuAffinity\" -Recurse
    Copy-Item ".\AutoGpuAffinity\config.ini" ".\build\AutoGpuAffinity\"

    return 0
}

$_exitCode = main
Write-Host # new line
exit $_exitCode
