@echo off
cls

cd "%~dp0"
pushd "%~dp0"

::PowerShell -NoProfile -STA -ExecutionPolicy Unrestricted -file "main.ps1"
PowerShell -NoProfile -STA -ExecutionPolicy RemoteSigned -file "main.ps1"

Pause