@echo off

setlocal ENABLEEXTENSIONS
set REG_NAME="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Misys Healthcare Systems\Misys Homecare"
set KEY_NAME=CachePath

FOR /F "usebackq skip=2 tokens=1-2*" %%A IN (`REG QUERY %REG_NAME% /v %KEY_NAME% 2^>nul`) DO ( set value=%%C)

::set folder="C:\Program Files (x86)\Allscripts Homecare\Client\Cache"
pushd "%value%"

tasklist /FI "IMAGENAME eq mhc.exe" 2>NUL | find /I /N "mhc.exe">NUL
if ERRORLEVEL 1 goto Process_NotFound

:Process_Found
::If Allscripts is running, we need to exit out
taskkill /f /im mhc.exe
::Delete Cache
for /F "delims=" %%i in ('dir /b') do (rmdir "%%i" /s/q || del "%%i" /s/q)

::Run reg tweak to reset toolbar
reg delete "HKCU\Software\Misys Healthcare Systems\Misys Homecare\Toolbar Settings" /f
goto END

:Process_NotFound
::Run reg tweak to reset toolbar
reg delete "HKCU\Software\Misys Healthcare Systems\Misys Homecare\Toolbar Settings" /f
::Delete Cache
for /F "delims=" %%i in ('dir /b') do (rmdir "%%i" /s/q || del "%%i" /s/q)
goto END

:END
::Run Allscripts again
::"C:\Program Files (x86)\Allscripts Homecare\Client\MHC.exe"
