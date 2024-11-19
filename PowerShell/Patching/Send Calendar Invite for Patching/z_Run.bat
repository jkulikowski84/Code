@echo off
cls

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0SendPatchingCalendarInvites.ps1'"

Pause
