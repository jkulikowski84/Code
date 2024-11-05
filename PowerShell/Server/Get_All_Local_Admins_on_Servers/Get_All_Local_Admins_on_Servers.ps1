CLS

#Load our modules
#List the modules we want to load per line
$LoadModules = @'
SplitPipeline
ImportExcel
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

#Clear Variables
Clear-Variable ADServers, DMZServers, ServersSkipCitrix, FilteredServers -Force -Confirm:$False -ErrorAction SilentlyContinue

#Get all Servers (filter out disabled computers)
$ADServers = ((dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*)(!(servicePrincipalName=*MSClusterVirtualServer*)))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }) | sort -Unique

#Set DMZ Credentials
if($NULL -eq $DMZCred)
{
    #Store DMZ credentials
    $DMZCred = Get-Credential( (whoami) -replace "domain","domain.local" )
}

#Get all DMZ Servers (filter out disabled computers)
$DMZServers = Invoke-Command -ComputerName "DMZ-AD.domain.local" -Credential $DMZCred -ScriptBlock { 
    (((dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*)(!(servicePrincipalName=*MSClusterVirtualServer*)))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }) | sort -Unique) <#|% { if($NULL -ne $_) { $("$_.domain.local") } }#>
} -ErrorAction SilentlyContinue

#=================================

$ServersSkipCitrix = foreach($item in $ADServers)
{
    if(($item -notlike "*ngld*") -AND ($item -notlike "*ecta*") -AND ($item -notlike "*ncta*"))
    {
        $item
    }
}

#=================================
$report_list = @()


$FilteredServers = $ServersSkipCitrix | split-pipeline -count 64 {
    Process
    {
        #Clear variables every iteration
        $ServerIP = $NULL

        #Get the IP from the hostname
        Try
        {
            $ServerIP = [System.Net.Dns]::GetHostAddresses("$($_)")
        }
        Catch {  Out-Null }

        Try
        {
            if(([System.Net.NetworkInformation.ping]::new().SendPingAsync($_).Wait(1000)) -eq "True")
            {
                $_
            }
        }
        Catch
        {
            if(([System.Net.NetworkInformation.ping]::new().SendPingAsync($ServerIP).Wait(1000)) -eq "True")
            {
                $_
            }
        }
    }
}

#=================================
#Get Server local admin accounts

$report_list = $FilteredServers | split-pipeline -count 64 {
    Process
    {
        function Get-LocalGroupMembers
        {
            $arr = @()

            Try
            {
                $wmi = Get-WmiObject -ErrorAction Stop -ComputerName $Server -Query "SELECT * FROM Win32_GroupUser WHERE GroupComponent=`"Win32_Group.Domain='$Server',Name='Administrators'`""
            }
            Catch
            {
                #Write-Host "$Server"
            }

            # Parse out the username from each result and append it to the array.
            if ($wmi -ne $null) 
            {
                foreach ($item in $wmi) 
                {
                    $arr += (($item.PartComponent.subString(($item.PartComponent.indexOf("Domain=") + 8), ($item.PartComponent.indexOf('",Name=') - ($item.PartComponent.indexOf("Domain=") + 8)))) + "\" + ($item.PartComponent.Substring($item.PartComponent.IndexOf(',') + 1).Replace('Name=', '').Replace("`"", '')))
                }
            }
            else 
            {
                $arr += "NULL"
            }

            $hash = @{ComputerName = $Server; GroupName = 'Administrators'; Members = $arr }
            return $hash
	
            end {}
        }

        $Server = $_

        Try
        {
            $Details = Get-LocalGroupMembers -ComputerName $Server -GroupName "Administrators"

            $GS_Group_Members = ""
            $SVC_Group_Members = ""
            $Non_GS_SVC_Group_Members = ""

            foreach ($Group_Member in $Details.Members)
            {
                if($Group_Member.subString($Group_Member.indexOf("\")+1) -Like "GS_*")
                {
                    $GS_Group_Members += $Group_Member
                
                    if($Group_Member -ne $Details.Members[-1])
                    {
                        $GS_Group_Members += ",`n"
                    }
                }
                elseif($Group_Member.subString($Group_Member.indexOf("\")+1) -Like "SVC_*")
                {
                    $SVC_Group_Members += $Group_Member

                    if($Group_Member -ne $Details.Members[-1])
                    {
                        $SVC_Group_Members += ",`n"
                    }
                }
                else
                {
                    $Non_GS_SVC_Group_Members += $Group_Member

                    if($Group_Member -ne $Details.Members[-1])
                    {
                        $Non_GS_SVC_Group_Members += ",`n"
                    }
                }            
            }

            $reportObject = New-Object System.Object
            $reportObject | Add-Member -MemberType NoteProperty -Name "GroupName" -Value "Administrators"
            $reportObject | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $Server
            $reportObject | Add-Member -MemberType NoteProperty -Name "GS_Members" -Value $GS_Group_Members.Trim()
            $reportObject | Add-Member -MemberType NoteProperty -Name "SVC_Members" -Value $SVC_Group_Members.Trim()
            $reportObject | Add-Member -MemberType NoteProperty -Name "Non_GS_SVC_Members" -Value $Non_GS_SVC_Group_Members.Trim()

            $reportObject
        }
        Catch
        {
            #$Server
        }
    }
}

$report_list | Export-Excel -Path "C:\Temp\LocalAdminUsers.xlsx" -WorksheetName "Server Local Admin" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"














<#

$report_list = $FilteredServers | Start-RSJob -Throttle 50 -Batch "Test" -ScriptBlock {
    Param($Server)

    function Get-LocalGroupMembers 
    {
        $arr = @()

        Try
        {
            $wmi = Get-WmiObject -ErrorAction Stop -ComputerName $Server -Query "SELECT * FROM Win32_GroupUser WHERE GroupComponent=`"Win32_Group.Domain='$Server',Name='Administrators'`""
        }
        Catch
        {
            Write-Host "$Server"
        }

        # Parse out the username from each result and append it to the array.
        if ($wmi -ne $null) 
        {
            foreach ($item in $wmi) 
            {
                $arr += (($item.PartComponent.subString(($item.PartComponent.indexOf("Domain=") + 8), ($item.PartComponent.indexOf('",Name=') - ($item.PartComponent.indexOf("Domain=") + 8)))) + "\" + ($item.PartComponent.Substring($item.PartComponent.IndexOf(',') + 1).Replace('Name=', '').Replace("`"", '')))
            }
        }
        else 
        {
            $arr += "NULL"
        }

        $hash = @{ComputerName = $Server; GroupName = 'Administrators'; Members = $arr }
        return $hash
	
        end {}
    }

    Try
    {
        $Details = Get-LocalGroupMembers -ComputerName $Server -GroupName "Administrators"

        $GS_Group_Members = ""
        $SVC_Group_Members = ""
        $Non_GS_SVC_Group_Members = ""

        foreach ($Group_Member in $Details.Members)
        {
            if($Group_Member.subString($Group_Member.indexOf("\")+1) -Like "GS_*")
            {
                $GS_Group_Members += $Group_Member
                
                if($Group_Member -ne $Details.Members[-1])
                {
                    $GS_Group_Members += ",`n"
                }
            }
            elseif($Group_Member.subString($Group_Member.indexOf("\")+1) -Like "SVC_*")
            {
                $SVC_Group_Members += $Group_Member

                if($Group_Member -ne $Details.Members[-1])
                {
                    $SVC_Group_Members += ",`n"
                }
            }
            else
            {
                $Non_GS_SVC_Group_Members += $Group_Member

                if($Group_Member -ne $Details.Members[-1])
                {
                    $Non_GS_SVC_Group_Members += ",`n"
                }
            }            
        }

        $reportObject = New-Object System.Object
        $reportObject | Add-Member -MemberType NoteProperty -Name "GroupName" -Value "Administrators"
        $reportObject | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $Server
        $reportObject | Add-Member -MemberType NoteProperty -Name "GS_Members" -Value $GS_Group_Members.Trim()
        $reportObject | Add-Member -MemberType NoteProperty -Name "SVC_Members" -Value $SVC_Group_Members.Trim()
        $reportObject | Add-Member -MemberType NoteProperty -Name "Non_GS_SVC_Members" -Value $Non_GS_SVC_Group_Members.Trim()

        $reportObject
    }
    Catch
    {
        #$Server
    }

} | Wait-RSJob -ShowProgress -Timeout 30 | Receive-RSJob

$report_list | Export-Excel -Path "C:\Temp\AAAAA.xlsx" -WorksheetName "Server Local Admin" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
#>