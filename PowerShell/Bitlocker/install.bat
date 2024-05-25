pushd %~dp0

::We need to make sure these Reg Keys exist and are set properly
reg add "HKLM\Software\Policies\Microsoft\FVE" /v ActiveDirectoryBackup /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v RequireActiveDirectoryBackup /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v ActiveDirectoryInfoToStore /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v OSActiveDirectoryBackup /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v OSActiveDirectoryInfoToStore /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v OSRequireActiveDirectoryBackup /t REG_DWORD /d 1 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v OSRecoveryKey /t REG_DWORD /d 2 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v OSRecoveryPassword /t REG_DWORD /d 2 /f
reg add "HKLM\Software\Policies\Microsoft\FVE" /v OSRecovery /t REG_DWORD /d 1 /f

::Now let's backup the recovery key to AD
PowerShell -NoProfile -STA -ExecutionPolicy Unrestricted -file "%~dp0Bitlocker.ps1"
Pause
