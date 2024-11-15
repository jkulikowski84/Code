CLS

$Users = Get-ADObject -LDAPFilter "(&(samAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(legacyExchangeDN=ADCDisabledMail))" -Properties *

#$Users | Export-Csv -NoTypeInformation -Path "C:\TEMP\legacyExchangeDN-USERS-All-Properties.csv"

 #$Result = foreach($Property in (($users[0]) | Get-Member))

foreach($Property in (($users) | Get-Member))
 {
    if($Property.MemberType -eq "Property")
    {
        if(($Property.Definition -like "*Int64*") -AND (($Property.Definition -notlike "*ADPropertyValueCollection*")))
        {
            if($($($users[0]).$($Property.Name)) -ne 0)
            {
                Write-Output "$($Property.Name): $([DateTime]::FromFileTime($($($users[0]).$($Property.Name))))"
            }
            else
            {
                Write-Output "$($Property.Name): $($($users[0]).$($Property.Name))"
            }
        }
        if((($Property.Definition -notlike "*Int64*") -AND (($Property.Definition -like "*ADPropertyValueCollection*"))) -AND ($Property.Definition -notlike "*WriteWarningStream*") -AND ((($Property.Definition -like "*ADPropertyValueCollection*"))))
        {
            if($NULL -ne $($($users[0]).$($Property.Name)))
            {
                #%{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }
                #Write-Output "$($Property.Name): $($($users[0]).$($Property.Name) )`n"
                if($Property.name -eq "memberOf")
                {
                    Write-Output "$($Property.Name):`n$($($($($users[0]).$($Property.Name))) |% { if($_ -match '^CN=(.+?),\s*\w{1,2}=') { $matches[1] } -join "`n"} )`n"
                }
                else
                {
                    Write-Output "$($Property.Name):`n$($($users[0]).$($Property.Name) |% { $_ -join "`n"} )`n"
                }
            }
        }
        else
        {
            if(($NULL -ne $($($users[0]).$($Property.Name))) -AND (($Property.Definition -notlike "*Int64*") -AND (($Property.Definition -notlike "*ADPropertyValueCollection*"))))
            {
                Write-Output "$($Property.Name): $($($users[0]).$($Property.Name) )`n"
            }
        }
    }
 }

 #$Result | Export-Csv -NoTypeInformation -Path "C:\TEMP\zzzzzz.csv"
