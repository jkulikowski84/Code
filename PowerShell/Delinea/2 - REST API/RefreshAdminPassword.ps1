CLS

#The admin account we want to get the password for
$SearchFilter = "adminAccount"

#Clear Variables each session
Clear-Variable session, antiForgeryToken, AdminUserID, Password, NewPW, Content, line, OldPW, NewContent, proc -Force -Confirm:$False -ErrorAction SilentlyContinue

#Reset Session
#$session = $NULL

$ihawu = "LongString"
$ThycoticLocation = "LongStringB"

#SessionCookie
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
$session.Cookies.Add((New-Object System.Net.Cookie("Thycotic_Location", "$ThycoticLocation", "/", "domain.secretservercloud.com")))
$session.Cookies.Add((New-Object System.Net.Cookie("ihawu", "$ihawu", "/", "domain.secretservercloud.com")))

#Grab the unique Token for our session
Do
{
	$antiForgeryToken = (Invoke-WebRequest -UseBasicParsing -Uri "https://domain.secretservercloud.com/internals/csrf" `
    -WebSession $session `
    -Headers @{
        "authority"="domain.secretservercloud.com"
        "method"="GET"
        "path"="/internals/csrf"
        "scheme"="https"
        "accept"="application/json, text/plain, */*"

    }).Content.Split('"')[3]

}While($NULL -eq $antiForgeryToken)

#Grab the userID of the account we are searching
Do
{
    $AdminUserID = ((Invoke-RestMethod -UseBasicParsing -Uri "https://domain.secretservercloud.com/api/v2/secrets?filter.doNotCalculateTotal=true&filter.includeActive=true&filter.includeRestricted=true&filter.permissionRequired=1&filter.scope=All&filter.searchText=$SearchFilter&skip=0&sortBy%5B0%5D.direction=asc&sortBy%5B0%5D.name=name&take=60" `
    -WebSession $session `
    -Headers @{
        "authority"="domain.secretservercloud.com"
        "method"="GET"
        "path"="/api/v2/secrets?filter.doNotCalculateTotal=true&filter.includeActive=true&filter.includeRestricted=true&filter.permissionRequired=1&filter.scope=All&filter.searchText=$SearchFilter&skip=0&sortBy%5B0%5D.direction=asc&sortBy%5B0%5D.name=name&take=60"
        "scheme"="https"
        "accept"="application/json"
        "x-requestverificationtoken"="$antiForgeryToken"
    }).records | Where-Object { $_.name -notlike "*extnch*"}).ID

}While($NULL -eq $AdminUserID)

#Get Password
Do
{
    $Password = (Invoke-RestMethod -UseBasicParsing -Uri "https://domain.secretservercloud.com/api/v1/secrets/$AdminUserID/fields/password?args.noAutoCheckout=true&args.includeInactive=true" `
    -WebSession $session `
    -Headers @{
        "authority"="domain.secretservercloud.com"
        "method"="GET"
        "path"="/api/v1/secrets/$AdminUserID/fields/password?args.noAutoCheckout=true&args.includeInactive=true"
        "scheme"="https"
        "accept"="application/json, text/plain, */*"
        "accept-encoding"="gzip, deflate, br, zstd"
    }).split('"')[1]

}While($NULL -eq $Password)

$NewPW = $Password

#=========== Get Old Admin PW from QuickTextPaste

#Location of the config file we want to modify
$QTP = "C:\Files\QuickTextPaste\QuickTextPaste.ini"

#Read the content of the file
$Content = [System.IO.File]::ReadAllLines($QTP)

#Get the line we want to work with
$line = $Content | Select-Object | Where-Object {$_ -like "*L-Win+E*"}

#Pull the current password from the file
$OldPW = (($line -split ('-p '))[1]).split(' ')[0]

#Replace Old Password with new one
$NewContent = $Content.Replace("$OldPW","$NewPW")

[System.IO.File]::WriteAllLines("$QTP", $NewContent, [System.Text.Encoding]::Unicode)

#Restart QTP Client
$proc = Get-Process -Name QuickTextPaste_x64_p | Sort-Object -Property ProcessName -Unique

if($NULL -ne $proc)
{
    $proc.Kill()
    Start-Sleep -s 1
}

#Restart our process
Invoke-Item "C:\Files\QuickTextPaste\QuickTextPaste_x64_p.exe"