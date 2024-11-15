CLS

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

#===================

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

#===================  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Load our Worksheet
Try
{
    $worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ErrorAction Stop -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name
}
Catch
{
    #Try loading our DLL and try again
    $DLL = "$((gmo -list importexcel).Path | split-path -Parent)\EPPlus.dll"
    Add-Type -Path $DLL
    $worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name
}

#$Servers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

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
                $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey(�Users�,"$Server")
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
                $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey(�LocalMachine�,"$Server")
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

#===================

$properties = @('distinguishedName', 'name', 'operatingSystem')

$Servers = $NULL

$Servers = (Get-ADObject -LDAPFilter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*)(!(servicePrincipalName=*MSClusterVirtualServer*)))" -Properties $properties)

$Exclusion = @("ECTA","NCTA","NGLD")
$OSCheck = @("2003","2008","2012")

$FilteredServers = ForEach ($Item in $Servers) 
{
    #Exclude servers that match the naming convention
    if(($Item.name -match ($Exclusion -join '|')) -eq $False)
    {
        if(($Item.operatingSystem -match ($OSCheck -join '|')) -eq $True)
        {
            $Item
        }
    }
}

#===================

$App = "Azure"

$FilteredServers | split-pipeline -count 64 -Function GetApp -Variable App -ErrorAction SilentlyContinue {
    
    Process
    {
        $Result = $Server = $NULL
        $Server = $_.Name

        $Result = GetApp -Server $Server -App $App

        if($NULL -eq $Result)
        {
            $Server
        }
    }
}


