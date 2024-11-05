This Installer/Uninstaller is custom built by me and I haven't come across any other script/package that does what this script does.

The Installer installs MS Teams as a Machine-Wide Install and then modified all the necessary data and registry hives so that it operated properly. If you want to uninstall it, you can easily do that like any other uninstall of an application.

The install and uninstall is extremely quick and takes less than a minute and I haven't come across any issues.

Things to keep in mind... 
For my purpose, we don't use Teams as our chat client. We only use it for meetings (mostly external meetings hosted by vendors). We don't have a MS tenant configured for Teams so I ran into the issue of waiting for team to stop authenticating or going in circles before I can get to it. I have disabled the annoyance by modifying the 2 JSON files that I am providing. The 2 JSON files will automatically be placed in the proper path during install, and since they are modified as Read Only, they won't be overwritten. 

What the above changes do is allow you to use Teams in Guest Mode without logging in or authenticating. If you do use MS Teams and have a tenant configured for it, then you probably don't need these altered JSON files, so you can delete them from your "%AppData%\Microsoft\Teams" path.

You can also download the latest MSI for MS Teams from:
https://teams.microsoft.com/desktopclient/installer/windows/x64 <-- Get the version from this link, and replace the extension from ".exe" to ".msi" (so it looks similar to the link below)

https://statics.teams.cdn.office.net/production-windows-x64/1.7.00.27855/Teams_windows_x64.msi (Replace the version with the newest one in this URL)

You can also get the versions from the links below:
https://github.com/ItzLevvie/MicrosoftTeams-msinternal/blob/master/defconfig2
https://stealthpuppy.com/apptracker/apps/m/microsoftteamsclassic/
