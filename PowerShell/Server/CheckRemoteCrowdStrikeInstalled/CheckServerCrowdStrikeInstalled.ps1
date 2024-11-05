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

Function CheckCrowdStrike($Server)
{
    $UninstallKey = ”SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall” 

    #Create an instance of the Registry Object and open the HKLM base key
    $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey(‘LocalMachine’,"$Server")

    #Drill down into the Uninstall key using the OpenSubKey Method
    $regkey = $reg.OpenSubKey($UninstallKey) 

    #Retrieve an array of string that contain all the subkey names
    $subkeys = $regkey.GetSubKeyNames() 

    foreach($key in $subkeys)
    {
        $thiskey = $UninstallKey + "\\" + $key
        $thisSubKey = $reg.OpenSubKey($thiskey)
        ($thisSubKey.GetValue("DisplayName") | Where-Object { $_ -like "*CrowdStrike*"})
    }
}

#grab all of our servers
$ADServers = ((Get-ADObject -LDAPFilter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*)(servicePrincipalName=*CmRcService*)(!(servicePrincipalName=*MSClusterVirtualServer*)))" -Properties dNSHostName ) | %{ if ($_ -notlike "*OU=Decom,OU=Servers,DC=domain,DC=com*"){ $_ } } ).dNSHostName | sort

#Exclude Citrix Servers
$AllNonCitrixServers = ($ADServers | Where-Object { ($_ -notlike "*ECTA*") -AND ($_ -notlike "*NCTA*") -AND ($_ -notlike "*NGLD*")}).replace(".domain.com","")

$AllNonCitrixServers | split-pipeline -count 64 -Function CheckCrowdStrike -ErrorAction SilentlyContinue {
    Process
    {
        $Test = $Server = $NULL

        $Server = $_

        Try
        {
            $Test = CheckCrowdStrike($Server)

            if($NULL -eq $Test)
            {
                $server
            }
        }
        Catch
        {
            $server
        }
    }
}
