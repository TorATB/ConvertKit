@echo off
echo Run as Administrator to uninstall
cd %APPDATA%\Microsoft\Windows\SendTo
del ConvertKit.lnk >nul 2>nul
cd %~dp0
cd..
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ConvertKit" /f
rmdir ConvertKit /s /q >nul 2>nul