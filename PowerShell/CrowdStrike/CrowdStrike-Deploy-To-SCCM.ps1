CLS

$LoadModules = @'
SplitPipeline
'@.Split("`n").Trim()

ForEach($Module in $LoadModules) 
{
    #Check if Module is Installed
    if($NULL -eq (Get-Module -ListAvailable $Module))
    {
        #Install Module
        Install-Module -Name $Module -Scope CurrentUser -Confirm:$False -Force
        #Import Module
        Import-Module $Module
    }

    #Check if Module is Imported
    if($NULL -eq (Get-Module -Name $Module))
    {
        #Install Module
        Import-Module $Module
    }
}

#===================

Function GetApp
{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Server,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $App
    )

    #Set variable
    $regPaths = @()

    if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node")
    {
        $regPaths += "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    }

    $regPaths += "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $regPaths += "*\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

    # Iterate through the paths
    foreach ($path in $regPaths) 
    {
        #Clear Variables
        $reg = $regkey = $subkeys = $NULL

        #If it's HKEY_Users
        if($path -like "*\SOFTWARE*")
        {
            Try
            {
                $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey(‘Users’,"$Server")
            }
            Catch
            {
                $Server
            }
        }
        else #It's HKEY_LOCAL_MACHINE
        {
            Try
            {
                $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey(‘LocalMachine’,"$Server")
            }
            Catch
            {
                $Server
            }
        }

        if($NULL -ne $reg)
        {
            #Drill down into the Uninstall key using the OpenSubKey Method
            $regkey = $reg.OpenSubKey($path) 

            #Retrieve an array of string that contain all the subkey names
            $subkeys = Try { $regkey.GetSubKeyNames() } Catch {}

            foreach($key in $subkeys)
            {
                #Clear Variables
                $thiskey = $thisSubKey = $NULL

                $thiskey = $path + "\\" + $key
                $thisSubKey = $reg.OpenSubKey($thiskey)
                ($thisSubKey.GetValue("DisplayName") | Where-Object { $_ -like "*$App*"})
            }
        }
    }
}

#=================== Remotely execute our Code through SCCM

$SCCMServer = "SCCMServer"
$SessionsRunning = get-pssession

#Create a new Session to the SCCM Server
if(!($SessionsRunning.ComputerName -like $SCCMServer) -and ($env:COMPUTERNAME -ne $SCCMServer))
{
    $Session = New-PSSession -ComputerName $SCCMServer -Authentication Kerberos
}

#====================================================================

#=================== Grab our Servers

#Clear Variables
Clear-Variable properties, data, ResponsiveServers, Servers, FilteredServers -Force -Confirm:$False -ErrorAction SilentlyContinue

$properties = @('distinguishedName', 'dNSHostName', 'name', 'operatingSystem')

$Servers = $NULL

$Servers = (Get-ADObject -LDAPFilter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*)(!(servicePrincipalName=*MSClusterVirtualServer*)))" -Properties $properties)

#Exclude Citrix
$Exclusion = @("ECTA","NCTA","NGLD")
#Exclude certain distinguishednames
$DNExclusion = @("Decom","Workstation")

$FilteredServers = ForEach ($Item in $Servers) 
{
    #Exclude servers that match the naming convention
    if(($Item.dNSHostName -match ($Exclusion -join '|')) -eq $False)
    {
        if(($Item.distinguishedName -match ($DNExclusion -join '|')) -eq $False)
        {
            $Item
        }
    }
}

#=================== Filter by responsive servers

#Used for progress
$data = @{
    Count = $(($FilteredServers | measure).count)
	Done = 0
}

$ResponsiveServers = $FilteredServers | split-pipeline -count 64 -Variable data -ErrorAction SilentlyContinue {

    Process
    {
        $Server = $_.name

        #Get the IP from the hostname
        Try
        {
            $ServerIP = [System.Net.Dns]::GetHostAddresses("$($Server)")
        }
        Catch
        {
            Try
            {
                $ServerIP = [System.Net.Dns]::GetHostAddresses("$((Get-ADObject -LDAPFilter "name=$Server" -Properties dNSHostName).dNSHostName)")
            }
            Catch
            {
                Write-Host "$($Server)" -ForegroundColor red -BackgroundColor white
            }
        }

        Try
        {
            #Ping by hostname
            if(([System.Net.NetworkInformation.ping]::new().SendPingAsync($Server).Wait(1000)) -eq "True")
            {
                $Server
            }
        }
        Catch
        {
            #Ping by IP
            if(([System.Net.NetworkInformation.ping]::new().SendPingAsync($ServerIP).Wait(1000)) -eq "True")
            {
                $Server
            }
        }

        $done = ++$data.Done

        # show progress
        Write-Progress -Activity "Done $done" -Status Processing -PercentComplete (100*$done/$data.Count)
    }
} | sort

Write-Progress -Activity 'Completed' -Completed

#=================== Search all of our responsive servers to see if they have crowdstrike

Clear-Variable data, NeedCrowdStrike -Force -Confirm:$False -ErrorAction SilentlyContinue

$App = "crowdstrike"

#Used for progress
$data = @{
    Count = $(($ResponsiveServers | measure).count)
	Done = 0
}

$NeedCrowdStrike = $ResponsiveServers | split-pipeline -count 64 -Function GetApp -Variable App, data -ErrorAction SilentlyContinue {
    
    Process
    {
        $Result = $Server = $NULL
        $Server = $_

        $Result = GetApp -Server $Server -App $App

        if(($NULL -eq $Result) -OR ($Result -like "*$Server*"))
        {
            $Server
        }

        $done = ++$data.Done

        # show progress
        Write-Progress -Activity "Done $done" -Status Processing -PercentComplete (100*$done/$data.Count)
    }
} | sort

Write-Progress -Activity 'Completed' -Completed

#=================== Execute the rest of our code within the context of our SCCM server

Invoke-Command -Session $Session -ScriptBlock {

    if($NULL -eq $SiteCodeObjs)
    {
        $SiteCodeObjs = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $env:COMPUTERNAME -ErrorAction Stop
    }

    if(($NULL -ne $SiteCodeObjs) -AND ($NULL -eq $SiteCode))
    {
        foreach ($SiteCodeObj in $SiteCodeObjs)
        {
            if ($SiteCodeObj.ProviderForLocalSite -eq $true)
            {
                $SiteCode = $SiteCodeObj.SiteCode
            }
        }
    }

    if(($NULL -ne $SiteCode) -AND ($NULL -eq $SitePath))
    {
        $SitePath = $SiteCode + ":"
    }

    if(((Get-Command -Module ConfigurationManager | Measure-Object).Count) -eq 0)
    {
        Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1')
    }

    #Make sure the Module & Drive are configured
    Do
    {
        $NULL = Get-Module configurationmanager | fl *
        Sleep -Milliseconds 100

    } While ($NULL -eq ((Get-PSDrive) | Select-Object | Where-Object { $_.Name -eq "$SiteCode" }))

    #Set our Patch to the CMSite
    if((Get-Location -ErrorAction SilentlyContinue).Path -ne $($SitePath + "\"))
    {
        Set-location $sitepath
    }

    ### ======== Start of main Code Block ========

    # Create new Device collection to deploy CrowdStrike

    $CollectionName = "CrowdStrike Server Deployment"
    $LimitingCollectionName = "Servers | All" # The collection to limit membership
    $Comment = "Server Collection used to push CrowdStrike to all of our servers"

    New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollectionName -Comment $Comment -ErrorAction SilentlyContinue

    #=================== 

    <#
        Instead of adding each device to the collection, we will write a query.
        Adding each device individually takes a REALLY long time
        Writing a custom query takes a second
    #>

    #First let's clear out any variables that might have old data
    $FinalQuery = $Query = $NULL

    #The start of our Query
    $Query = "SELECT SMS_R_SYSTEM.Name from SMS_R_System where SMS_R_System.Name in ("

    #We will Append each server into the query string
    foreach ($server in $using:NeedCrowdStrike)
    {
        $Query += '"' + $server + '",'
    }

    #Clean up the Query
    $FinalQuery = $Query.TrimEnd(",") + ")"

    #===================

    #The name of our collection is above (line 281). Let's get the Collection Info.
    $collection = Get-CMDeviceCollection -Name $CollectionName

    # Add our Query to the membership rule
    Add-CMDeviceCollectionQueryMembershipRule -CollectionId $collection.CollectionID -QueryExpression ($FinalQuery) -RuleName "Test"

    # Update the collection membership
    Invoke-CMCollectionUpdate -Name $CollectionName
}
