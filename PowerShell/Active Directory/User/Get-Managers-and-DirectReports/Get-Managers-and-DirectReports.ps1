CLS

#Get All Managers
$Managers = (dsquery * -filter "(&(objectClass=person)(objectCategory=Person)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" -limit 0 -attr Manager | sort -Unique ).trim()

#Get all users they manage
Foreach($Manager in $Managers)
{
    Try
    {
        Write-Output "Manager: $((Get-ADUser $B).GivenName + " " + (Get-ADUser $B).Surname)"
        Write-Output " "
        (Get-ADUser $B -Properties * | select @{Name="DirectReports";Expression={($_.directreports | %{ (Get-ADUser $_).GivenName + " " + (Get-ADUser $_).Surname })}}).DirectReports | Sort
        Write-Output " "
        Write-Output ("=" * 40)
        Write-Output " "
    }
    Catch
    {
    
    }
}