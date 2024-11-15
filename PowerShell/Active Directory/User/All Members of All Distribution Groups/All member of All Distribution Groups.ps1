CLS

$DistributionGroups = Get-ADGroup -Filter {GroupCategory -eq "Distribution"}

($DistributionGroups.Name) | sort
<#
foreach($DistributionGroup in $DistributionGroups)
{
    Write-Output "`t$($DistributionGroup.Name)`n"

    ((dsquery * -filter "(&(memberof=$($DistributionGroup.DistinguishedName)))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } })

    Write-Output "`n"
}
#>