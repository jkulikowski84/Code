CLS

$RandomIPs = 1..254 | % { [IPAddress]::Parse([String] (Get-Random) ).IPAddressToString }

$RandomIPs | Sort-Object { [System.Version]($_) }

