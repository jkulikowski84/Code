CLS

#SRC: https://blog.idera.com/database-tools/final-super-fast-ping-command

#Untested: https://www.powershellgallery.com/packages/Soap/5.1.4/Content/Scripts%5CPing-Sweep.ps1
#https://gist.github.com/TheRockStarDBA/6cdd505500ccd81a7c326bf0d5991f26

#Variables
$TimeoutMillisec = 1000

#Ask user for an IP
Do
{
    CLS
    $IP = Read-Host "Type in an IP"

}until($IP -match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$" -and [bool]($IP -as [ipaddress]))

#Strip the last octet of he IP and create a range to scan from the IP above
$IPRange = 1..254 | % {"$($IP.Substring(0, $IP.lastIndexOf('.'))).$_"}

[Collections.ArrayList]$bucket = @()

$StatusCode_ReturnValue = 
        @{
            0='Success'
            11001='Buffer Too Small'
            11002='Destination Net Unreachable'
            11003='Destination Host Unreachable'
            11004='Destination Protocol Unreachable'
            11005='Destination Port Unreachable'
            11006='No Resources'
            11007='Bad Option'
            11008='Hardware Error'
            11009='Packet Too Big'
            11010='Request Timed Out'
            11011='Bad Request'
            11012='Bad Route'
            11013='TimeToLive Expired Transit'
            11014='TimeToLive Expired Reassembly'
            11015='Parameter Problem'
            11016='Source Quench'
            11017='Option Too Big'
            11018='Bad Destination'
            11032='Negotiating IPSEC'
            11050='General Failure'
        }

$statusFriendlyText = @{
            # name of column
            Name = 'Status'
            # code to calculate content of column
            Expression = { 
                # take status code and use it as index into
                # the hash table with friendly names
                # make sure the key is of same data type (int)
                $StatusCode_ReturnValue[([int]$_.StatusCode)]
            }
        }

        # calculated property that returns $true when status -eq 0
        $IsOnline = @{
            Name = 'Online'
            Expression = { $_.StatusCode -eq 0 }
        }

        # do DNS resolution when system responds to ping
        $DNSName = @{
            Name = 'DNSName'
            Expression = { if ($_.StatusCode -eq 0) { 
                    if ($_.Address -like '*.*.*.*') 
                    { [Net.DNS]::GetHostByAddress($_.Address).HostName  } 
                    else  
                    { [Net.DNS]::GetHostByName($_.Address).HostName  } 
                }
            }
        }

$IPRange | ForEach-Object {
            $null = $bucket.Add($_)
        }

$query = $bucket -join "' or Address='"

(Get-WmiObject -Class Win32_PingStatus -Filter "(Address='$query') and timeout=$TimeoutMillisec" | Select-Object -Property Address, $IsOnline, $DNSName, $statusFriendlyText ) | where-object { ($_.Status -ne "Success") } | Sort-Object { [System.Version]($_.Address) }