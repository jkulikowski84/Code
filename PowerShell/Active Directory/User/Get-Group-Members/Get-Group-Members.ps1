CLS

$DynamicGroup = "Nursing Excellence"

#Get Current Users in the group
$CurrentGroupMembers = dsquery * -filter "(&(memberof=CN=$DynamicGroup,OU=Security Groups,OU=domain,DC=domain,DC=com))" -limit 0 | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }

$CurrentGroupMembers | sort 