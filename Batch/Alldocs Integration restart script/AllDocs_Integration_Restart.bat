@ECHO OFF
CLS

::If Hyland is running in the background, close the process
tasklist /FI "IMAGENAME eq Hyland.Applications.Allscripts.exe" 2>NUL | find /I /N "Hyland.Applications.Allscripts.exe">NUL
if ERRORLEVEL 0 ( taskkill /f /im Hyland.Applications.Allscripts.exe > NUL 2>&1 )

start "" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Integration for Allscripts.lnk"
