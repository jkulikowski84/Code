@echo off
cls

echo Installing Files. Please be patient.
@echo off
cls

pushd "%~dp0"

PowerShell -NoProfile -STA -ExecutionPolicy Unrestricted -file "%~dp0Shared-Mailbox.ps1"

Pause
