CLS

$Users = Get-ADObject -LDAPFilter "(&(samAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(legacyExchangeDN=ADCDisabledMail))" -Properties DisplayName,DistinguishedName,givenname,legacyExchangeDN,Name,ObjectClass,ObjectGUID,sn,mail

$Users | Export-Csv -NoTypeInformation -Path "C:\TEMP\legacyExchangeDN-USERS.csv"
