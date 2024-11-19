CLS

#------------- Install VMware PowerCLI if it doesn't already exist

if((Get-Module -ListAvailable -Name "VMware.PowerCLI") -or (Get-Module -Name "VMware.PowerCLI"))
{
    #Import-Module VMware.PowerCLI
}
else
{	
    Install-Module -Name VMware.PowerCLI -Scope CurrentUser -Force -AllowClobber -Confirm:$False
	Import-Module VMware.PowerCLI
}

#------------- PoshRSJob (multitasking)

if((Get-Module -ListAvailable -Name "PoshRSJob") -or (Get-Module -Name "PoshRSJob"))
{
        Import-Module PoshRSJob
}
else
{	
    Install-Module -Name PoshRSJob -Scope CurrentUser -Force -Confirm:$False
	Import-Module PoshRSJob
}


#------------- ImportExcel Module 

if((Get-Module -ListAvailable -Name "ImportExcel") -or (Get-Module -Name "ImportExcel"))
{
        Import-Module ImportExcel
}
else
{
    #Install NuGet (Prerequisite) first
	Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$False
	
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$False
	Import-Module ImportExcel
}

#Clear screen again
CLS

#----------------------------------------------------------------------------------------------------------------

#Start Timestamp
$Start = Get-Date

#Global Variables
$Path = (Split-Path $script:MyInvocation.MyCommand.Path)
$ErrorFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\ERROR.csv"

$Results = "$Path\Results.xlsx"

if((Test-Path $Results) -eq $True)
{
    Remove-Item $Results
}

if((Test-Path $ErrorFile) -eq $True)
{
    Remove-Item $ErrorFile
}

#------------------------------------  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

$ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

#------------------------------------ Populate our variable with data from spreadsheet

$ExcelServersList = foreach($ExcelServer in $ExcelServers) 
{
    $ExcelServer | Select-Object @{Name="ServerName";Expression={$_.Child}}, "Primary", @{Name="PatchWindow";Expression={$_."Patch Window"}}, @{Name="TestServer";Expression={$_."Test Server"}}, "DMZ", "Operating System"
}

#[Net.ServicePointManager]::SecurityProtocol = 'Tls', 'Tls11','Tls12'

#Set-PowerCLIConfiguration -Scope User -DefaultVIServerMode Multiple -ParticipateInCEIP $false -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $False -Confirm:$False | Out-NULL

if($NULL -eq ($global:DefaultVIServers.Name))
{
    #$cred = (Get-Credential (whoami))
    
    #Connect-VIServer "pvsa02vcsa0001" -Protocol https -Credential $cred -AllLinked -WarningAction 0 | Out-NULL
    Connect-VIServer "pvsa02vcsa0001" -Protocol https -Credential (Import-clixml "$Path\vSphere.clixml") -AllLinked -WarningAction 0 | Out-NULL
}

#------------------------------------

$VS_Servers = (Get-VM) -Join '|'

$sView = @{

   Filter = @{"Name" = [string]$($VS_Servers.Name); 'Config.Template' = 'false'; "Config.GuestFullName" = "Microsoft" }

   ViewType = 'VirtualMachine'

   Property = 'Name', 'RunTime.BootTime'
}

$VS_View = Get-View @sView | Select Name, @{N="LastBoot";E={($_.Runtime.Boottime).tolocaltime()}} | Sort-Object -Property Name

#------------------------------------

Function GetPhysicalServers
{
    $PhysicalServers = @()

    ForEach ($Item in $ExcelServersList)
    {
        If($item.DMZ -eq $True)
        {
            #Remove DMZ suffix from name in order to compare with vsphere
            $item.servername = $item.servername -replace (".dmz.com","")
        }

        If($item.servername -notin $VS_View.Name)
        {
            $PhysicalServers += $item
        }

        If($item.DMZ -eq $True)
        {
            #Readd the DMZ suffix
            $item.servername = [System.String]::Concat("$($item.ServerName)",".dmz.com")
        }
    }

    return $PhysicalServers
}

Function GetVirtualServers
{
    $VirtualServers = @()

    ForEach ($Item in $ExcelServersList)
    {
        If($item.DMZ -eq $True)
        {
            #Remove DMZ suffix from name in order to compare with vsphere
            $item.servername = $item.servername -replace (".dmz.com","")
        }

        If($item.servername -in $VS_View.Name)
        {
            $VirtualServers += $item
        }
    }

    return $VirtualServers
}

Function GetPoweredOffServers
{
    $GetPoweredOffServers = @()

    ForEach ($ItemA in $ExcelServersList)
    {
        If($ItemA.DMZ -eq $True)
        {
            #Remove DMZ suffix from name in order to compare with vsphere
            $ItemA.servername = $ItemA.servername -replace (".dmz.com","")
        }

        ForEach ($ItemB in $VS_View)
        {
            If(($itemA.servername -in $ItemB.Name) -AND ($NULL -eq $itemB.LastBoot))
            {
                $GetPoweredOffServers += $itemA
            }
        }
    }

    return $GetPoweredOffServers
}

Function GetVirtualServerReboots
{
    $VirtualServerReboots = @()

    ForEach ($ItemA in $ExcelServersList)
    {
        If($ItemA.DMZ -eq $True)
        {
            #Remove DMZ suffix from name in order to compare with vsphere
            $ItemA.servername = $ItemA.servername -replace (".dmz.com","")
        }

        ForEach ($ItemB in $VS_View)
        {
            If(($ItemA.servername -in $ItemB.Name) -AND ($NULL -ne $itemB.LastBoot))
            {
                $Check = ((Get-Date) - ($itemB.LastBoot)).Days
                
                if($Check -ge 25)
                {
                    $VirtualServerReboots += $ItemA | Add-Member -NotePropertyMembers @{"Day's Online" = $Check} -PassThru -Force
                }
            }
        }
    }

    return $VirtualServerReboots
}

Function GetPhysicalServerReboots
{
    $PhysicalServers = GetPhysicalServers

    $PhysicalServerReboots = @()

    ForEach ($ItemA in $PhysicalServers)
    {
        Try
        {
            #$ServerPath = "\\$($Server.ServerName)\c$\Windows\ServiceProfiles\LocalService\NTUSER.DAT"
            $ServerPath = "\\$($ItemA.ServerName)\c$\Windows\ServiceProfiles\LocalService"

            #If you're checking file date
            #$Check = ((Get-Date) - (dir $ServerPath -force -ErrorAction Stop).LastWriteTime).days
            $Check = ((Get-Date) - ((dir $ServerPath -force -ErrorAction Stop -Filter "NTUSER*") | sort LastWriteTime -Descending).LastWriteTime[0]).Days
        }
        Catch
        {
            Try
            {
                $Check = ((Get-Date) - (Get-WmiObject win32_operatingsystem -ComputerName $($ItemA.ServerName) -ErrorAction Stop | select csname, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}).LastBootUpTime).days
            }
            Catch
            {                
                if($Error.exception -like "*")
                {
                    ($ItemA | Add-Member -NotePropertyMembers @{"Error" = [string]$Error} -PassThru) | Export-Csv -Path $ErrorFile -NoTypeInformation -Force -Append
                }

                #$Error | Select Property *
            }
        }
        
        #Check if server hasn't been rebooted in greater than or equal to 25 days.
        if($Check -ge 25)
        {
            $PhysicalServerReboots += $ItemA | Add-Member -NotePropertyMembers @{"Day's Online" = $Check} -PassThru -Force
        }
    }

    return $PhysicalServerReboots
}

<#
$a = GetVirtualServerReboots
$b = GetPhysicalServerReboots
$c = ($a + $b) | sort $_.Servername

$c | Export-Excel -Path $Results -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
#>

(GetVirtualServerReboots + GetPhysicalServerReboots) | sort $_.Servername | Export-Excel -Path $Results -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#GetPoweredOffServers

#------------------------------------ Multithreading Magic

<#
$FilteredServersResult | Start-RSJob -Throttle 25 -ScriptBlock {
    Param($Server)

    #Ping servers to make sure they're responsive
    if($NULL -ne (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$($Server.servername)' AND Timeout=100").ResponseTime)
    { 
        Try
        {
            #$ServerPath = "\\$($Server.ServerName)\c$\Windows\ServiceProfiles\LocalService\NTUSER.DAT"
            $ServerPath = "\\$($Server.ServerName)\c$\Windows\ServiceProfiles\LocalService"

            #If you're checking file date
            #$Check = ((Get-Date) - (dir $ServerPath -force -ErrorAction Stop).LastWriteTime).days
            $Check = ((Get-Date) - ((dir $ServerPath -force -ErrorAction Stop -Filter "NTUSER*") | sort LastWriteTime -Descending).LastWriteTime[0]).Days
        }
        Catch
        {
            Try
            {
                $Check = ((Get-Date) - (Get-WmiObject win32_operatingsystem -ComputerName $($Server.ServerName) -ErrorAction Stop | select csname, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}).LastBootUpTime).days
            }
            Catch
            {                
                if($Error.exception -like "*")
                {
                    ($Server | Add-Member -NotePropertyMembers @{"Error" = [string]$Error} -PassThru) | Export-Csv -Path $using:ErrorFile -NoTypeInformation -Force -Append
                }

                #$Error | Select Property *
            }
        }
        
        #Check if server hasn't been rebooted in greater than or equal to 25 days.
        if($Check -ge 25)
        {
            $Server | Add-Member -NotePropertyMembers @{"Day's Online" = $Check} -PassThru
        }
    }
} | Wait-RSJob -ShowProgress -Timeout 30 | Receive-RSJob | Export-Excel -Path $Results -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
#} | Wait-RSJob -ShowProgress | Receive-RSJob | Export-Csv -Path "$Path\Results.csv" -NoTypeInformation -Force 
#>
$End =  (Get-Date)

$End - $Start
