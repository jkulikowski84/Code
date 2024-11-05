CLS

# Gets todays Date
$date = Get-Date

# Number of days it's been since the computer authenticated to the domain
# In my case 1 day
$days = "-90"

# This is the OU you are searching for Stale Computer accounts
$ou = "OU=COMPUTERS,OU=COMPUTER-SYSTEMS,DC=domain,DC=com"

# Finding Stale Computers
$findcomputers = Get-adcomputer –filter * -SearchBase $ou -properties cn, LastLogonDate | 
Where {$_.LastLogonDate –le [DateTime]::Today.AddDays($days) -and ($_.lastlogondate -ne $null) }

# Create a CSV containg all the Stale Computer Information
#$findcomputers.Name

$findcomputers | Remove-ADObject -Recursive -Confirm:$false
