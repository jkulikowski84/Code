This file script checks to see if your current bitlocker key exists in AD.
If it exists in AD, it checks if it's the same key as what's locally on your machine.
If it matches, then it doesn't make any changes.
If the keys don't match, a new one is generated and logged in AD.

If it doesn't log properly in AD, you need to check the below registry path:

"HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE"
Check the following Keys:

All keys are DWORD 32
ActiveDirectoryBackup value set to Decimal 1
RequireActiveDirectoryBackup value set to Decimal 1
ActiveDirectoryInfoToStore value set to Decimal 1
OSActiveDirectoryBackup value set to Decimal 1
OSActiveDirectoryInfoToStore value set to Decimal 1
OSRequireActiveDirectoryBackup value set to Decimal 1
OSRecoveryKey value set to Decimal 2
OSRecoveryPassword value set to Decimal 2
OSRecovery value set to Decimal 1
