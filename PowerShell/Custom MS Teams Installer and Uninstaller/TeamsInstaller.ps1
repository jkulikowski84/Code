CLS

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

#Get logged on users SID & username
$usersid = (New-Object -ComObject Microsoft.DiskQuota).TranslateLogonNameToSID((Get-WmiObject -Class Win32_ComputerSystem).Username)
$username = ((Get-WmiObject -Class Win32_ComputerSystem).Username).split("\")[1]

#Delete Reg key if it exists
if(Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams")
{
    if([bool]((Get-Item -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams" -ErrorAction SilentlyContinue ).GetValueNames() -eq "AllUser") -eq $True)
    {
        Remove-Item -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams" -ErrorAction SilentlyContinue
    }
}

#Add registry key to trick it into installing the Machine-Wide Teams Installation
if([bool](Get-Item -LiteralPath "HKLM:\SOFTWARE\Citrix\PortICA" -ErrorAction SilentlyContinue ) -eq $False) { New-Item -Path "HKLM:\SOFTWARE\Citrix\PortICA" -Force | Out-Null }

#Install Teams
#Start-Process msiexec.exe "/i $Path\Teams_windows_x64.msi /l*v $Path\logfile.txt ALLUSER=1 ALLUSERS=1"
Start-Process -FilePath "$env:systemroot\system32\msiexec.exe" -ArgumentList "/i", "$Path\Teams_windows_x64.msi", "/l*v", "$Path\logfile.txt", 'ALLUSER="1"', 'ALLUSERS="1"' -Wait -NoNewWindow

if([bool](Get-Item -LiteralPath "HKLM:\SOFTWARE\Citrix\PortICA" -ErrorAction SilentlyContinue ) -eq $True) { Remove-Item -LiteralPath "HKLM:\SOFTWARE\Citrix\PortICA" -ErrorAction SilentlyContinue }

#We installed the Machine Wide Installer. Now for future updates to work, we need to change the key.
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams" -Name "AllUser" -PropertyType String -Value 0 -Force | Out-Null

#Normally users don't have permission over Program Files. Let's give them the necessary permissions so they can update.
Start-Process -FilePath "$Path\SetACL.exe" -ArgumentList '-on "C:\Program Files (x86)\Microsoft\Teams"', '-ot "file"', '-actn "ace"', '-ace "n:users;p:Change"' -Wait -NoNewWindow | Out-Null

#Create a directory for our uninstaller if it doesn't already exist
if(!(Test-Path -Path "C:\Users\$($username)\AppData\Local\TeamsUninstaller")) { New-Item -Path "C:\Users\$($username)\AppData\Local\TeamsUninstaller" -ItemType Directory | Out-Null }

#Copy our uninstaller
Copy-Item -Path "$Path\Uninstaller\*" -Destination "C:\Users\$($username)\AppData\Local\TeamsUninstaller" -Force

#Rename our shortcut
Rename-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic.lnk" -NewName "Microsoft Teams.lnk"

<#
NOTES
Microsoft Teams Installs 2 copies of the Application.
Teams_Installer, and then the actual application.
Once we have the APplication installed, we don't need to keep the installer, however if you uninstall the Installer, it will also uninstall your full App.
Instead, we just hide the App Installer. The App will still uninstall through Programs and Features.
#>

New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -Name "DisplayIcon" -PropertyType String -Value 'C:\Program Files (x86)\Microsoft\Teams\current\Teams.exe,0' -Force | Out-Null
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -Name "DisplayName" -PropertyType String -Value 'Microsoft Teams' -Force | Out-Null
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -Name "UnInstallString" -PropertyType String -Value "C:\Users\$($username)\AppData\Local\TeamsUninstaller\Uninstall.bat" -Force | Out-Null
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -Name "ModifyPath" -PropertyType String -Value '' -Force | Out-Null
Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -Name "WindowsInstaller" -Force | Out-Null

<#
This path stores the "Installer", but when we do an update, this path doesn't get overwritten, because the update writes to a different location.
To prevent this from happening we copy this key to the location where the update installs so it will properly overwrite the values.
This way we only see 1 version of Teams like we should.
NOTE - when this script is ran as an admin then HKCU is in the escalated users context instead of the user's computer.
To get the proper users context we need to get the user that's logged on
#>

#Copy registry from old destination to users HIVE
Copy-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -Destination "registry::HKEY_USERS\$($usersid)\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams" -Recurse -Force -Confirm:$False

#Delete old location
if([bool](Get-Item -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -ErrorAction SilentlyContinue ) -eq $True) { Remove-Item -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}" -ErrorAction SilentlyContinue }

#Now delete the Installer because we no longer need it.
if(Test-Path("C:\Program Files (x86)\Teams Installer"))
{
    Remove-Item -Path "C:\Program Files (x86)\Teams Installer" -Recurse -Force
}

#Delete the packages folder. We don't need it.
if(Test-Path("C:\Program Files (x86)\Microsoft\Teams\packages"))
{
    Remove-Item -Path "C:\Program Files (x86)\Microsoft\Teams\packages" -Recurse -Force
}

#Create a shortcut on public Desktop
Copy-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk" -Destination "$env:Public\Desktop" -Force -ErrorAction SilentlyContinue | Out-Null

#Stop msiexec if it's still running
Stop-Process -Name "Msiexec" -Force -ErrorAction SilentlyContinue

#Copy over a basic desktop-config.json file that prevents Teams from trying to authenticate/log you in.
Copy-Item -Path "$Path\desktop-config.json" -Destination "C:\Users\$username\AppData\Roaming\Microsoft\Teams" -Force -ErrorAction SilentlyContinue | Out-Null
Copy-Item -Path "$Path\settings.json" -Destination "C:\Users\$username\AppData\Roaming\Microsoft\Teams" -Force -ErrorAction SilentlyContinue | Out-Null
