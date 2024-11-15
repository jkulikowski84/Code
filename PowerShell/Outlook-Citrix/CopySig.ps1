CLS

$HomeDrive = "\\CitrixProfiles\home$\$($env:username)"

#Check which folder exists
if(test-path("$($HomeDrive)\AppData\Roaming\Microsoft\Signatures"))
{
    $Source = "$($HomeDrive)\AppData\Roaming\Microsoft\Signatures"
}

if(test-path("$($HomeDrive)\Application Data\Roaming\Microsoft\Signatures"))
{
    $Source = "$($HomeDrive)\Application Data\Roaming\Microsoft\Signatures"
}

if(test-path("$($HomeDrive)\Application Data\Microsoft\Signatures"))
{
    $Source = "$($HomeDrive)\Application Data\Microsoft\Signatures"
}

$Destination = "\\CitrixServer\c$\Users\$($env:username)\AppData\Roaming\Microsoft"

Copy-Item -Path $Source -Destination $Destination -Recurse:$True -Confirm:$False -Force
