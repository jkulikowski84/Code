@echo off
cls

pushd "%~dp0"

PowerShell -NoProfile -STA -ExecutionPolicy Unrestricted -file "%~dp0CreateFieldShortcut.ps1"