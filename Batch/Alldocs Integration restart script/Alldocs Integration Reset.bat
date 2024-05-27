@echo off

tasklist /FI "IMAGENAME eq mhc.exe" 2>NUL | find /I /N "mhc.exe">NUL
if ERRORLEVEL 1 goto Process_NotFound

:Process_Found
::If Allscripts is running, we need to exit out
taskkill /f /im mhc.exe > NUL 2>&1

::Close out of alldocs
taskkill /f /im Hyland.Applications.Allscripts.exe > NUL 2>&1

::Restart Alldocs
start "" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Integration for Allscripts.lnk"

::wait 5 seconds before launching Allscripts (to make sure alldocs is running)
ping 127.0.0.1 -n 5 > NUL 2>&1

::Restart Allscripts
::start "" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Allscripts Healthcare\Allscripts Homecare Client.lnk"
start "" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Netsmart Technologies\Netsmart Homecare Client.lnk"
goto END

:Process_NotFound
::Close out of alldocs
taskkill /f /im Hyland.Applications.Allscripts.exe > NUL 2>&1

::Restart Alldocs
start "" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Integration for Allscripts.lnk"

::wait 5 seconds before launching Allscripts (to make sure alldocs is running)
ping 127.0.0.1 -n 5 > NUL 2>&1

::Restart Allscripts
::start "" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Allscripts Healthcare\Allscripts Homecare Client.lnk"
start "" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Netsmart Technologies\Netsmart Homecare Client.lnk"
:END

