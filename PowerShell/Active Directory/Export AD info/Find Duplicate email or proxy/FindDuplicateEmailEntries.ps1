CLS

function Get-EmailAddress
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True,
				   ValueFromPipeline = $True,
				   ValueFromPipelineByPropertyName = $True,
				   HelpMessage = 'What e-mail address would you like to find?')]
		[string[]]$EmailAddress
	)
	
	process
	{		
		foreach ($address in $EmailAddress)
		{
			Get-ADObject -Properties mail, proxyAddresses -Filter "mail -like '*$address*' -or proxyAddresses -like '*$address*'"
		}
	}
}

$EmailAddy = "email@domain.com"
Get-EmailAddress $EmailAddy