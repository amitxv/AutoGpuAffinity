function main() {
    if (Test-Path ".\build\") {
        Remove-Item -Path ".\build\" -Recurse
    }

    $entryPoint = "..\..\AutoGpuAffinity\main.py"

    # create folder structure
    New-Item -ItemType Directory -Path ".\build\AutoGpuAffinity\"

    # pack executable
    New-Item -ItemType Directory -Path ".\build\pyinstaller\"
    Push-Location ".\build\pyinstaller\"
    pyinstaller $entryPoint --onefile --name AutoGpuAffinity
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
