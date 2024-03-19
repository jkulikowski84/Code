schtasks.exe /Run /TN "\Microsoft\Windows\Servicing\StartComponentCleanup"
Dism.exe /online /Cleanup-Image /SPSuperseded
::c:\windows\SYSTEM32\cleanmgr.exe /dDrive
c:\windows\SYSTEM32\cleanmgr.exe