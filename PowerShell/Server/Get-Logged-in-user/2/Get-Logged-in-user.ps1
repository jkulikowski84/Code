CLS

#Get logged on users SID
$usersid = (New-Object -ComObject Microsoft.DiskQuota).TranslateLogonNameToSID((Get-WmiObject -Class Win32_ComputerSystem).Username)
$usersid
$username = ((Get-WmiObject -Class Win32_ComputerSystem).Username).split("\")[1]
$username
