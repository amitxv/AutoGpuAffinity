function Is-Admin() {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function main() {
    if (-not (Is-Admin)) {
        Write-Host "error: administrator privileges required"
        return 1
    }

    if (Test-Path ".\tmp\") {
        Remove-Item -Path ".\tmp\" -Recurse -Force
    }

    mkdir ".\tmp\"

    $urls = @{
        "liblava"           = "https://github.com/liblava/liblava.git"
        "presentmon_1.6.0"  = "https://github.com/GameTechDev/PresentMon/releases/download/v1.6.0/PresentMon-1.6.0-x64.exe"
        "presentmon_latest" = "https://github.com/GameTechDev/PresentMon/releases/download/v1.10.0/PresentMon-1.10.0-x64.exe"
        "CRU"               = "https://www.monitortests.com/download/cru/cru-1.5.2.zip"
        "D3D9_benchmark"    = "https://github.com/amitxv/Benchmark-DirectX9/releases/download/v2.0/Benchmark.DirectX9.Black.White.exe"
    }

    if (Test-Path ".\AutoGpuAffinity\bin\") {
        Remove-Item -Path ".\AutoGpuAffinity\bin\" -Recurse -Force
    }

    # create bin folder
    mkdir ".\AutoGpuAffinity\bin\"

    # =============
    # Setup liblava
    # =============

    git clone $urls["liblava"] ".\tmp\liblava\"
    Push-Location ".\tmp\liblava\"

    # build binaries
    cmake -B build -DCMAKE_BUILD_TYPE=Release -DCMAKE_INSTALL_PREFIX=install
    cmake --build build --config Release

    Pop-Location

    # copy binary to bin directory
    mkdir ".\AutoGpuAffinity\bin\liblava\"
    Copy-Item ".\tmp\liblava\build\Release\lava-triangle.exe" ".\AutoGpuAffinity\bin\liblava\"

    # ================
    # Setup PresentMon
    # ================

    mkdir ".\AutoGpuAffinity\bin\PresentMon\"

    Invoke-WebRequest $urls["presentmon_1.6.0"] -OutFile ".\AutoGpuAffinity\bin\PresentMon\PresentMon-1.6.0-x64.exe"
    Invoke-WebRequest $urls["presentmon_latest"] -OutFile ".\AutoGpuAffinity\bin\PresentMon\PresentMon-1.10.0-x64.exe"

    # ===============
    # Setup restart64
    # ===============

    mkdir ".\AutoGpuAffinity\bin\restart64\"

    Invoke-WebRequest $urls["CRU"] -OutFile ".\tmp\CRU.zip"
    Expand-Archive -Path ".\tmp\CRU.zip" -DestinationPath ".\tmp\CRU\"
    Copy-Item ".\tmp\CRU\restart64.exe" ".\AutoGpuAffinity\bin\restart64\"
    Copy-Item ".\tmp\CRU\Info.txt" ".\AutoGpuAffinity\bin\restart64\LICENSE.txt"

    # ====================
    # Setup D3D9 benchmark
    # ====================

    Invoke-WebRequest $urls["D3D9_benchmark"] -OutFile ".\AutoGpuAffinity\bin\D3D9-benchmark.exe"

    return 0
}

$_exitCode = main
Write-Host # new line
exit $_exitCode
