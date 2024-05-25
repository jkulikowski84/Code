CLS

#-----------------------------------------
#          Load proper Assemblies 
#-----------------------------------------

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#-----------------------------------------
#             Global Variables 
#-----------------------------------------

$global:File = $null
$global:text = $null
$global:FullFilePath = $null
$global:adminCreds = $null
$global:FileName = $null
$global:ProcessName = $null
$global:InstallerRan = $false

$global:adminCreds = ""

$monitor = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
$global:path = (Split-Path $script:MyInvocation.MyCommand.Path)
$global:Computer = $env:COMPUTERNAME

#-----------------------------------------
#     Pre-configured MSI batch script 
#-----------------------------------------

$msiBatch = @"
@echo off
pushd "%~dp0"

::Get the MSI file name to install
for /f "tokens=1* delims=\" %%A in ( 'forfiles /s /m *.msi /c "cmd /c echo @relpath"' ) do for %%F in (^"%%B) do (set myapp=%%~F)

::Launch our installer
start /w "" msiexec /i "%~dp0%myapp%" /passive /norestart

::Self Delete
DEL "%~f0"
"@

function Authentication
{
	try
	{
		#---------------------------------------------------
		#Authenticate Admin Account using encrypted password
		#---------------------------------------------------
		
		#Temp folder that will have the files
		$TempFolder = $env:temp
		
		#The Credentials are stored in 2 encrypted files
		#$global:AESKeyFilePath = $path + "\aeskey.txt"
		#$global:SecurePwdFilePath =  $path + "\credpassword.txt"
		$global:AESKeyFilePath = $TempFolder + "\aeskey.txt"
		$global:SecurePwdFilePath =  $TempFolder + "\credpassword.txt"		
		$global:userUPN = "journeycare\scheduler"

		#use key to create local secure passwordtemp
		$global:AESKey = Get-Content -Path $AESKeyFilePath -Force
		$global:securePass = Get-Content -Path $SecurePwdFilePath -Force | ConvertTo-SecureString -Key $AESKey

		#create a new psCredential object with required username and password
        $global:adminCreds = New-Object System.Management.Automation.PSCredential($userUPN, $securePass)		 
	}
	Catch [Exception]
	{
		$ErrorMessage = $_.Exception.Message
	}
	return $ErrorMessage
}

function Cleanup
{
    #---------------------------------------------------
    #      Cleanup the files that are extracted
    #---------------------------------------------------

	try
	{
		CLS
		$CleanupScript = $path + "\RemoveFiles.bat"
		#Start-Process -FilePath $CleanupScript -Credential $adminCreds -NoNewWindow -ArgumentList "-command &{Start-Process $CleanupScript -verb runas}" -Wait -WorkingDirectory $path -ErrorAction SilentlyContinue
		Start-Process -FilePath $CleanupScript -Credential $adminCreds -WindowStyle Hidden -ArgumentList "-command &{Start-Process $CleanupScript -verb runas}" -PassThru -Wait -WorkingDirectory $path -ErrorAction SilentlyContinue
		CLS
	}
	Catch [Exception]
	{
		$ErrorMessage = $_.Exception.Message
	}
	return $ErrorMessage
}

function RemoveAccount
{
	try
	{
		#Variables
		$name = $userUPN.split('\')[-1]
		$AccountSID = (New-Object System.Security.Principal.NTAccount($name)).Translate([System.Security.Principal.SecurityIdentifier]).value
		$LocalAccount = (gwmi Win32_UserProfile | Where-Object {$_.SID -eq $AccountSID})
		$UserAccount = $LocalAccount.LocalPath.split('\')[-1]
		
		Get-CimInstance -Class Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object {$_.LocalPath -like $name} | Remove-CimInstance -ErrorAction SilentlyContinue
	}
	Catch [Exception]
	{
		$ErrorMessage = $_.Exception.Message
	}
	return $ErrorMessage	
}

function GetParameters
{
    #---------------------------------------------------
    #      This is the input dialog that comes up
    #---------------------------------------------------

	try
	{
		Add-Type -AssemblyName Microsoft.VisualBasic

		$title = 'Application Parameters'
		$msg   = 'Type in the parameters you want to run your file with
		   Wait for the dialog with parameters to come up 
		Example: /s /v /qn'
		$DefaultValue = ""
		$XPos = ((($monitor.width) / 2) - 180)
		$YPos = ((($monitor.height) / 2) - 300)
		
		$text = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title, $DefaultValue, $XPos, $YPos)
		
		#Close the process we opened up earlier with the parameters since we don't need it anymore
		(gwmi Win32_Process -ComputerName $computer | ?{ $_.ProcessName -match $ProcessName }).Terminate() | Out-Null

		#Now launch the application with the parameters we passed through it
		Start-Process powershell.exe -LoadUserProfile -Credential $adminCreds -WindowStyle Hidden -Args '-command', "Start-Process $FullFilePath -Verb RunAs -Args '$text'" -PassThru -Wait
	}
	Catch [Exception]
	{
		$ErrorMessage = $_.Exception.Message
	}
	return $ErrorMessage
}

function BrowseToFile
{
	try
	{
		Add-Type -AssemblyName System.windows.forms

		$BrowseToFile = New-Object System.Windows.Forms.OpenFileDialog
		$BrowseToFile.Title = "Choose the file you want to run with escalated permissions"
		$BrowseToFile.initialDirectory = $path
		$BrowseToFile.ShowDialog() | Out-Null
		$global:File = $BrowseToFile.filename

		#===============================================

		#After the user selects the Application they want to run
		#Launch the dialog with the applications command line parameters
		#So we can customize the installation

		$FullFilePath = $global:File
		$FileName = Split-Path $FullFilePath -leaf
		$ProcessName = [io.path]::GetFileNameWithoutExtension($FullFilePath)

		if($FileName -like "*.bat" -or $FileName -like "*.cmd")
		{
			#Append a few lines to the top of our index file so that it doesn't prompt any "false" errors to users
			@("@echo off","cls","","echo Installing Files. Please be patient.") +  (Get-Content $FullFilePath) | Set-Content $FullFilePath
			
			#Now launch the installer
			Start-Process -FilePath $FullFilePath -LoadUserProfile -Credential $adminCreds -NoNewWindow -ArgumentList "-command &{Start-Process $FullFilePath -verb runas}" -Wait -WorkingDirectory $path				
			
			#Remove the changes we made to the file
			(Get-Content $FullFilePath | Select-Object -Skip 4) | Set-Content $FullFilePath
		}
		if($FileName -like "*.msi")
		{
			$installPath = $path + "\msiInstaller.bat"
			$msiBatch | Out-File -Encoding Ascii -append $installPath
			$FullFilePath = $installPath
			
			Start-Process -FilePath $FullFilePath -LoadUserProfile -Credential $adminCreds -NoNewWindow -ArgumentList "-command &{Start-Process $FullFilePath -verb runas}" -Wait -WorkingDirectory $path	
		}
		else
		{		
			#Display the App Command line Parameters
			Start-Process powershell.exe -Credential $adminCreds -Args '-noprofile', '-command', "Start-Process '$FullFilePath' -Verb RunAs -Args '/?'"

			do{
				#This is just a "holder" to check when our app window comes up so we can bring up the input dialog.
			}until((Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) |  where { $_.mainwindowhandle -ne 0 } )

			GetParameters
		}
	}
	Catch [Exception]
	{
		$ErrorMessage = $_.Exception.Message
	}
	return $ErrorMessage
}

function Run-SoftwareEscalated
{
    Try 
    {       
        #Run our installer
		#This array will hold the common names and setup packages of software setups
		[array]$installer = "install.bat", "install.cmd", "*.msi"
        
		for($i=0; $i -lt $installer.Count; $i++)
		{
			$filter = Get-ChildItem $path -Filter $($installer[$i]) -name
			$FullFilePath = $path + "\" + $filter
			
			#Make sure the file exists. If it doesn't then skip it
			if(Test-Path $FullFilePath -PathType Leaf)
			{
				if($filter -like "*.bat" -or $filter -like "*.cmd")
				{
					#Append a few lines to the top of our index file so that it doesn't prompt any "false" errors to users
					@("@echo off","cls","","echo Installing Files. Please be patient.") +  (Get-Content $FullFilePath) | Set-Content $FullFilePath
					
					#Continue
				}
				if($filter -like "*.msi")
				{
					$installPath = $path + "\msiInstaller.bat"
					$msiBatch | Out-File -Encoding Ascii -append $installPath
					$FullFilePath = $installPath
				}
	
				#Start-Process -FilePath $FullFilePath -Credential $adminCreds -LoadUserProfile -NoNewWindow -ArgumentList "-command &{Start-Process $FullFilePath -verb runas}" -PassThru -WorkingDirectory $path                 
                Start-Process -FilePath $FullFilePath -LoadUserProfile -Credential $adminCreds -NoNewWindow -ArgumentList "-command &{Start-Process $FullFilePath -verb runas}" -PassThru -Wait -WorkingDirectory $path
				$InstallerRan = $true
                
				#Remove the top 4 lines in the batch script that we appended above
				if($FullFilePath -like "*.bat" -or $FullFilePath -like "*.cmd")
				{
					(Get-Content $FullFilePath | Select-Object -Skip 4) | Set-Content $FullFilePath
				}
				break
			}
		}
		#If no batch scripts or MSI packages are found, prompt user to point to the path of the file they want to install.
		if($InstallerRan -ne $true)
		{
			BrowseToFile
		}
    }
    Catch 
    {
        Write-Warning -Message "$($_.Exception.Message)"
    }
}

#Run the authentication function to authenticate our session
Authentication

#Launch our main application
Run-SoftwareEscalated

#Remove account after everything is complete
RemoveAccount

#Remove our unneeded files, because they were already loaded in
Cleanup