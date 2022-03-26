@echo off
pyinstaller AutoGpuAffinity.py --onefile --add-binary "lava-triangle.exe;." --add-binary "PresentMon.exe;." --add-data "res.zip;." --uac-admin
exit /b