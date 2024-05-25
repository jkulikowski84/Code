@echo off

tasklist /FI "IMAGENAME eq mhc.exe" 2>NUL | find /I /N "mhc.exe">NUL
if ERRORLEVEL 1 goto Process_NotFound

:Process_Found
::If Allscripts is running, we need to exit out
taskkill /f /im mhc.exe

::Run reg tweak to reset toolbar
reg delete "HKCU\Software\Misys Healthcare Systems\Misys Homecare" /f
goto END

:Process_NotFound
::Run reg tweak to reset toolbar
reg delete "HKCU\Software\Misys Healthcare Systems\Misys Homecare" /f
goto END

:END

