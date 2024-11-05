CLS

$DaysInactive = 90

$time = (Get-Date).Adddays(-($DaysInactive))

$Stales = Get-ADObject -Filter {LastLogonTimeStamp -lt $time} -SearchBase "OU=COMPUTERS,OU=COMPUTER-SYSTEMS,DC=journeycare,DC=net"

foreach ($stale in $stales)
{
    $Stale #| Remove-ADObject -Recursive -Confirm:$false
    #Remove-ADComputer -Identity $Stale -Confirm:$False
}
