CLS

#GLOBAL VARIABLES
$global:Portal = $PortalSelection = $NULL

function SetPortal
{
    ##Variables
    $FilePath = "C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client"
    [array]$Files="config.ini","defaults.ini"

    #-------------------------------
    #If the PrinterLogic client is running, gracefully shut it down

    $PrinterInstallerClient = Get-Process PrinterInstallerClient -ErrorAction SilentlyContinue

    if ($PrinterInstallerClient)
    {
        #Start-Process -FilePath "C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\PrinterInstallerClient.exe" -ArgumentList "Shutdown"
		Start-Process -FilePath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Printer Installer\Administration\Shutdown Client.lnk"
		Start-Sleep -m 500
    }

    #-------------------------------
    #Check if the file we want to modify is read only

    foreach ($File in $Files)
    {
        $status = Get-ChildItem "$FilePath\$File"

        If ($status.isreadonly)
        {
            #If it's read only, change it so we can modify it
            $status.set_isreadonly($false)
        }

        #-------------------------------
        #Replace the text we want in the files

        #Set path first
        $Path = "$FilePath\$File"

        #Do our magic
        (Get-Content -Path "$Path" -Raw).ToLower() | Foreach-Object {
        $_ -replace 'printerportal.domain.com', $Portal `
           -replace 'printers.domain.com', $Portal `
           -replace 'prodprinter.domain.com', $Portal `
           -replace 'testPrinter.domain.com', $Portal
        } | Set-Content $Path

        #-------------------------------
        #After we're done modifying, set the file to read only so it doesn't get tampered with
        
        $status = Get-ChildItem "$FilePath\$File"

        If (!($status.isreadonly))
        {
            #If it's read only, change it so we can modify it
            $status.set_isreadonly($true)
        }
    }

    #Launch our client again
    #Start-Process -FilePath "C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\PrinterInstallerClient.exe" -ArgumentList "hideclientwindow"
	Start-Process -FilePath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Printer Installer\Administration\Start Client.lnk"
}

function Main
{
    Write-Host "Choose which PrinterPortal you want to set as your default"
    Write-Host ""
    Write-Host "A. printers.domain.com"
    Write-Host "B. prodprinter.domain.com\printerportal.domain.com"
    Write-Host "C. testPrinter.domain.com"
    Write-Host ""

    $PortalSelection = Read-Host "Press the letter corresponding to the PrinterPortal you want to set as your default or type 'exit' to exit"

    switch($PortalSelection)
    {
        A 
        {
            $Portal = "printers.domain.com"
            SetPortal
        }
        B 
        {
            $Portal = "printerportal.domain.com"
            SetPortal
        }
        C
        {
            $Portal = "testPrinter.domain.com"
            SetPortal
        }
        Exit 
        { 
            break 
        }
        Default 
        { 
            CLS
            Main
        }
    }
}

Main
