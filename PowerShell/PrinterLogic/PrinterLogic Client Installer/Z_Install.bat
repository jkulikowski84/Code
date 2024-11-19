@echo off
cls

::Uninstal Old Printerlogic Client
@echo off
cls

pushd "%~dp0"

taskkill /F /IM PrinterInstallerClientLauncher.exe
taskkill /F /IM PrinterInstallerClient.exe
taskkill /F /IM PrinterInstallerClientInterface.exe
taskkill /F /IM PrinterInstaller_MicrosoftMigrator.exe
net stop spooler
%windir%\system32\msiexec.exe /qn /quiet /norestart /x {A9DE0858-9DDD-4E1B-B041-C2AA90DCBF74} REMOVE=ALL

if exist "C:\Program Files (x86)\Printer Properties Pro" ( Del "C:\Program Files (x86)\Printer Properties Pro\*.*" /q )

reg delete HKLM\SOFTWARE\PrinterLogic /f
reg delete HKLM\SOFTWARE\Wow6432Node\PPP /f

if exist "%windir%\Temp\data" ( RMDIR %windir%\Temp\data /q /s )
if exist "%windir%\Temp\PPP" ( RMDIR %windir%\Temp\PPP /q /s )
if exist "C:\Program Files (x86)\Printer Properties Pro" (  RMDIR "C:\Program Files (x86)\Printer Properties Pro" /s /q )
net start spooler

::--------------------
CLS

::Install newest Client
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Install-PrinterLogicClient.ps1'
