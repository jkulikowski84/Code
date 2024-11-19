@echo off
cls

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Software-list-verifications.ps1'"

Pause
