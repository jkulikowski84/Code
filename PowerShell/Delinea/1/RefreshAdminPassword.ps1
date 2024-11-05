CLS

if((Get-Module -ListAvailable -Name "Thycotic.SecretServer") -or (Get-Module -Name "Thycotic.SecretServer"))
{
        #Import-Module Thycotic.SecretServer
}
else
{	
    Install-Module -Name Thycotic.SecretServer -RequiredVersion 0.61.0 -Scope CurrentUser -Force -Confirm:$False
	Import-Module Thycotic.SecretServer
}

if($NULL -eq $session)
{
    $token = "longencryptionkey"
    
    Try
    {
        $session = New-TssSession https://thycotic.domain.com/SecretServer -AccessToken $token -WarningAction Stop
    }
    Catch
    {
        #Write-Output "Token expired. Replace your token."
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageIcon = [System.Windows.MessageBoxImage]::Information
        $MessageBody = "Token expired. Replace your token."
        $MessageTitle = "Token Expired"
        $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

        if($NULL -ne $Result)
        {
            Exit
        }
    }
}

#=========== Get New Admin PW from Thycotic

$SecretID = (Thycotic.SecretServer\Find-TssSecret -TssSession $session -SearchText "SomeAccountPassWord" | Select-Object | Where-Object { $_.SecretName -notlike "*dmz*"}).SecretId

$NewPW = ((Get-TssSecret -TssSession $session -Id $SecretID).items | where-object {$_.FieldName -eq "Password"}).ItemValue

#=========== Get Old Admin PW from QuickTextPaste

#Location of the config file we want to modify
$QTP = "C:\Files\QuickTextPaste\QuickTextPaste.ini"

#Read the content of the file
$Content = [System.IO.File]::ReadAllLines($QTP)

#Get the line we want to work with
$line = $Content | Select-Object | Where-Object {$_ -like "*L-Win+E*"}

#Pull the current password from the file
$OldPW = (($line -split ('-p '))[1]).split(' ')[0]

#Replace Old Password with new one
$NewContent = $Content.Replace("$OldPW","$NewPW")

[System.IO.File]::WriteAllLines("$QTP", $NewContent, [System.Text.Encoding]::Unicode)

#Restart QTP Client
$proc = Get-Process -Name QuickTextPaste_x64_p | Sort-Object -Property ProcessName -Unique

if($NULL -ne $proc)
{
    $proc.Kill()
    Start-Sleep -s 1
}

#Restart our process
Invoke-Item "C:\Files\QuickTextPaste\QuickTextPaste_x64_p.exe"
