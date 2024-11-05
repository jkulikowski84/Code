function Get-ModuleAD() 
{
    #Add the import and snapin in order to perform AD functions
    #Get Primary DNS
	$DNS = (Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IpEnabled='True'" | ForEach-Object {$_.DNSServerSearchOrder})[1]

    #Convert IP to hostname
    $hostname = ([System.Net.Dns]::gethostentry($DNS)).HostName

    #Add the necessary modules from the server
    $session = New-PSSession -ComputerName $hostname -Authentication Kerberos
    Invoke-Command -Session $session -ScriptBlock {Get-Module ActiveDirectory} 
    Import-Module -Name ActiveDirectory -PSSession $session
}

#Load Active Directory Module remotely if it's not already loaded
if(!(Get-Module -ListAvailable -Name "ActiveDirectory"))
{
    Get-ModuleAD
}

#Clear screen
CLS

#Global Variables
$global:SystemDrive = [System.Environment]::GetLogicalDrives()[0]
$global:DriveLetter = $SystemDrive.TrimEnd('\')
$global:OSversion = [Environment]::OSVersion.Version -ge (new-object 'Version' 6,2)
$global:BitLocker = Get-WmiObject -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" -Filter "DriveLetter = '$DriveLetter'"
$global:ProtectorIds = $BitLocker.GetKeyProtectors().volumekeyprotectorID
$global:KeyProtectorType = $NULL
$global:Tpm = Get-WmiObject -Namespace ROOT\CIMV2\Security\MicrosoftTpm -Class Win32_Tpm
$LogFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Bitlocker-Log.txt"

#Uncomment the lines below to enable logging
#Write-Output "Computer Bitlocker Info for: $env:COMPUTERNAME" | Out-File -Append -FilePath $LogFile
#Write-Output "" | Out-File -Append -FilePath $LogFile

#This function checks to see if the recovery key in AD matches the recovery key set locally on the machine
Function MatchingKey
{
    #This checks what the local recovery key is
    foreach ($ProtectorID in $ProtectorIds)
    {
        $KeyProtectorType = $BitLocker.GetKeyProtectorType($ProtectorID).KeyProtectorType

        #If keytype = 3, that means it has a numerical password which is what we want
        if($KeyProtectorType -eq "3")
        {
            $KeyProtectorReturn = $BitLocker.GetKeyProtectorNumericalPassword($ProtectorID)
            $LocalRecoveryPassword = $KeyProtectorReturn.NumericalPassword
            
            #Uncomment the line below to enable logging
            #Write-Output "Local Recovery Key: $LocalRecoveryPassword" | Out-File -Append -FilePath $LogFile
            Write-Output "Local Recovery Key: $LocalRecoveryPassword"
        }
    }

    #This checks what recovery password is set in AD
    $computer = Get-ADComputer -Filter "Name -like '$env:computername'"
    $recoverykey = Get-ADObject -Filter 'objectclass -like "msFVE-RecoveryInformation"' -SearchBase $computer.DistinguishedName -Properties *
    #$ADRecoveryPassword = ($recoverykey.'msFVE-RecoveryPassword')
    $ADRecoveryPassword = ($recoverykey.'msFVE-RecoveryPassword') | Select-Object -Last 1

    #Uncomment the lines below to enable logging
    #Write-Output "AD Recovery Key: $ADRecoveryPassword" | Out-File -Append -FilePath $LogFile
    #Write-Output "" | Out-File -Append -FilePath $LogFile
    Write-Output "AD Recovery Key: $ADRecoveryPassword"

    #Let's make sure they are equal
    if($ADRecoveryPassword -eq $LocalRecoveryPassword)
    {
        #return $True
        #If the key matches we don't need to do anything
    }
    else
    {
        #return $False
        #If the key doesn't match, we need to reconfigure it.
               
        ####Some output for debugging
        #$ProtectorID
        #$KeyProtectorType
        #$LocalRecoveryPassword

        Backup-BitLockerKeyProtector -MountPoint $DriveLetter -KeyProtectorId $ProtectorID
    }
}

#Main Function
MatchingKey
#$CheckKey = MatchingKey #This variable stores either true or false; whether or not the recovery key matches or exists in AD. Even if it exists, it makes sure it's the correct one.

#Write-Output "Does Recovery Key Match what's in AD: $CheckKey" | Out-File -Append -FilePath $LogFile
#Write-Output "" | Out-File -Append -FilePath $LogFile

###DEBUG###
#Check TPM Info
#Uncomment the lineS below to enable logging
#Write-Output "Is tpm activated: $($tpm.IsActivated().isactivated)" | Out-File -Append -FilePath $LogFile
#Write-Output "Is tpm enabled: $($tpm.IsEnabled().isenabled)" | Out-File -Append -FilePath $LogFile
#Write-Output "Is ownership allowed: $($Tpm.IsOwnershipAllowed().IsOwnerShipAllowed)" | Out-File -Append -FilePath $LogFile
#Write-Output "Does TPM have an owner: $($Tpm.isowned().isowned)" | Out-File -Append -FilePath $LogFile
#Write-Output "---------------------------------------------" | Out-File -Append -FilePath $LogFile
#Write-Output "" | Out-File -Append -FilePath $LogFile

#Close out our session once we're done using it
if($null -ne $session)
{
    Remove-PSSession -Session $session
}
