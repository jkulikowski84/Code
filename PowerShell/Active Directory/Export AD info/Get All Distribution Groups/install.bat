@echo off
cls

pushd "%~dp0"

PowerShell -NoProfile -STA -ExecutionPolicy Unrestricted -file "%~dp0Export-All-Distribution-Groups-and-fields.ps1"
