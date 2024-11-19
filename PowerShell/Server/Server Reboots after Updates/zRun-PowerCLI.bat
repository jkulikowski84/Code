@echo off
cls

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0powerCLI.ps1'"

Pause
