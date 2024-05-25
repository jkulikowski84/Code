@echo off
cls

Powershell -NoProfile -STA -ExecutionPolicy Unrestricted -file "%~dp0alldocs.ps1"
