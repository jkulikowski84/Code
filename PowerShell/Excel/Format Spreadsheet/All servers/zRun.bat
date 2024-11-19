@echo off
cls

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Format-Spreadsheet.ps1'"

Pause
