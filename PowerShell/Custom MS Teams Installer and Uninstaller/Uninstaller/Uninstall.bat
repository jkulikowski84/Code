@echo off
cls

if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0TeamsUninstaller.ps1'"

exit