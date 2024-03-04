function main() {
    # pack executable
    pyinstaller AutoGpuAffinity\main.py --onefile --name AutoGpuAffinity

    # create folder structure
    mkdir building\AutoGpuAffinity
    Move-Item dist\AutoGpuAffinity.exe building\AutoGpuAffinity
    Move-Item AutoGpuAffinity\bin building\AutoGpuAffinity
    Move-Item AutoGpuAffinity\config.ini building\AutoGpuAffinity

    return 0
}

$_exitCode = main
Write-Host # new line
exit $_exitCode
