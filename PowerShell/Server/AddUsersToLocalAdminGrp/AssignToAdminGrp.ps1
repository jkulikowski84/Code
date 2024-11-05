CLS

$CopyServer = "TWVS01SPKW0011"
$NewServer = "TWVS01SPKA0012"

$LocalAdminUsers = $NULL

$LocalAdminUsers = Invoke-Command -ComputerName $CopyServer -ScriptBlock {
	$members = Invoke-Expression -command "Net Localgroup Administrators"
	$members[6..($members.Length-3)]
}

Invoke-Command -ComputerName $NewServer -ScriptBlock {
    
	foreach($user in $using:LocalAdminUsers)
	{
		Invoke-Expression -command "Net Localgroup Administrators /add '$($user)'" -ErrorAction SilentlyContinue
		Invoke-Expression -command "Add-LocalGroupMember -Group Administrators -member '$($user)'" -ErrorAction SilentlyContinue
	}
} -ErrorAction SilentlyContinue