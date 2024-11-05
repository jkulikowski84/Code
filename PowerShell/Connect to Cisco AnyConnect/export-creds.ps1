CLS

$Credential = Get-Credential
$Credential | Export-Clixml "C:\Scripts\VPN-Connection\creds.xml"
