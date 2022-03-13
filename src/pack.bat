@echo off
pyinstaller AutoGpuAffinity.py --onefile --add-data "lava-triangle.exe;." --add-data "PresentMon.exe;." --add-data "res.zip;."
exit /b