@echo off
cls

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Server-Reboot-Check.ps1'"

Pause
