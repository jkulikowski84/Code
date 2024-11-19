@echo off
cls

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Refresh-Configs.ps1'"

Pause
