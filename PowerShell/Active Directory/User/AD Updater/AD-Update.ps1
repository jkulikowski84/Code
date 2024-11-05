$OldPref = $global:ErrorActionPreference
$global:ErrorActionPreference = 'Stop'

function Get-ModuleAD() 
{
    #Add the import and snapin in order to perform AD functions
    #Get Primary DNS
    $DNS = (Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IpEnabled='True'" | ForEach-Object {$_.DNSServerSearchOrder})[1]

    #Convert IP to hostname
    $hostname = ([System.Net.Dns]::gethostentry($DNS)).HostName

    #Add the necessary modules from the server
    Try
    {
        $session = New-PSSession -ComputerName $hostname -Authentication Kerberos -ErrorAction Stop
    }
    Catch
    {
        $Auth = Get-Credential
        $session = New-PSSession -ComputerName $hostname -Authentication $Auth
    }

    Invoke-Command -Session $session -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $session
}

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

#Install the module that will let us perform certain tasks in Excel
#Install PSExcel Module for powershell
if(!(Get-Module -ListAvailable -Name ImportExcel) )
{
	#Install NuGet (Prerequisite) first
	Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$False
	
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$False
	Import-Module ImportExcel
} 

#Clear screen again
CLS

#----------------------------------------------------------------------------------------------------------------
 <#
    The worksheet variable will need to be modified before running this script. 
    Whatever the name of the worksheetis that you want to import data from, type that in below.
#>
$worksheet = "Sheet1"

#The file we will be reading from
$ExcelFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Book.xlsx"
$ErrorFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\ERROR.txt"

$Import = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

foreach ($User in $Import)
{
    $DisplayName = $User."AD Names"
    $Email = $User."Email Address"
    $Title = $User."HR Title"
    $tmobile = $User."T-Mobile Number"
    $fuze = $User."Fuze Phone Number"
    $BYODPhone = $User."BYOD Phone"
    $Supervisor = $User."HR Supervisor AD Name"
    $Department = $User."Team Name"

    Try
    {
        #$Manager = Get-ADUser -Filter { DisplayName -like $Supervisor }
		$Manager = Get-ADUser -Filter "DisplayName -like '$Supervisor'"
        $ManagerUN = $Manager.SamAccountName
    }
    Catch
    {
        Write-Output "Couldn't find Manager: $Supervisor for user: $DisplayName" | Out-File -Append -FilePath $ErrorFile
    }

    Try
    {
        #Validate the user exists in AD
        $validatedUser = Get-ADUser -Filter { UserPrincipalName -like $Email }
    }
    Catch
    {
        #We failed, lets get some information so we can find out why...
        Write-Output "Can't find user: $DisplayName" | Out-File -Append -FilePath $ErrorFile
    }

    #We will use a SID as the identifier for users. This is the most accurate method
    $SID = $validatedUser.SID.Value

    Try
    {
        #Clear out the fields we will be overwriting
        set-aduser $SID -Clear Manager, Department, Description, Title, HomePhone, telephoneNumber, mobile, ipPhone #-WhatIf

        #Populate fields
        Set-ADUSer -Identity $SID -Department $Department -Manager $ManagerUN -Description $Title -Title $Title

        #Since not everyone has a mobile number or phone number,
        #We only run it for users that have a phone on the spreadsheet

        if(([string]::IsNullOrWhiteSpace($BYODPhone)) -ne $True)
        {
            Try
            {
                #Try replacing the Phone Number
                Set-ADUser -Identity $SID -Replace @{ Mobile = $BYODPhone } #-WhatIf
            }
            Catch
            {
                #If the above fails, this is a failsafe
                Set-ADUSer -Identity $SID -MobilePhone $BYODPhone
            }
        }

        #if($NULL -ne $tmobile)
        if(([string]::IsNullOrWhiteSpace($tmobile)) -ne $True)
        {
            Try
            {
                #Try replacing the Phone Number
                Set-ADUser -Identity $SID -Replace @{ Mobile = $tmobile } #-WhatIf
            }
            Catch
            {
                #If the above fails, this is a failsafe
                Set-ADUSer -Identity $SID -MobilePhone $tmobile
            }
        }

        #if($NULL -ne $fuze)
        if(([string]::IsNullOrWhiteSpace($fuze)) -ne $True)
        {
            Try
            {
                #Try replacing the Phone Number
                Set-ADUser -Identity $SID -Replace @{ OfficePhone = $fuze } #-WhatIf
            }
            Catch
            {
                #If the above fails, this is a failsafe
                Set-ADUSer -Identity $SID -OfficePhone $fuze
            }
        }
    }
    Catch
    {
        Write-Output "Error for: $DisplayName" | Out-File -Append -FilePath $ErrorFile
        $_.Exception.Message | Out-File -Append -FilePath $ErrorFile
        $_.Exception.ItemName | Out-File -Append -FilePath $ErrorFile
        $_.InvocationInfo.MyCommand.Name | Out-File -Append -FilePath $ErrorFile
        $_.ErrorDetails.Message | Out-File -Append -FilePath $ErrorFile
        $_.InvocationInfo.PositionMessage | Out-File -Append -FilePath $ErrorFile
        $_.CategoryInfo.ToString() | Out-File -Append -FilePath $ErrorFile
        $_.FullyQualifiedErrorId | Out-File -Append -FilePath $ErrorFile
        Write-Output "-----------------------------" | Out-File -Append -FilePath $ErrorFile
        Write-Output "" | Out-File -Append -FilePath $ErrorFile
    }
}

#Close out our session once we're done using it
if($null -ne $session)
{
    Remove-PSSession -Session $session
}
