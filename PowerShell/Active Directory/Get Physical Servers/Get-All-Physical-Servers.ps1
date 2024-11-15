CLS

#=================== PoshRSJob (multitasking)

if((Get-Module -ListAvailable -Name "PoshRSJob"))
{
        #Import-Module PoshRSJob
}
else
{	
    Install-Module -Name PoshRSJob -Scope CurrentUser -Force -Confirm:$False
	Import-Module PoshRSJob
}

#Start Timestamp
$Start = (Get-Date)

#Clear Variables
Clear-Variable Servers, PhysicalServers, Check -Force -Confirm:$False -ErrorAction SilentlyContinue

$Servers = (dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0 -attr Name | sort).trim()

$PhysicalServers = $Servers | Start-RSJob -Throttle 50 -Batch "Test" -ScriptBlock {
    Param($Server)

    Try
    {
        $Check = [system.Net.Sockets.TcpClient]::new().BeginConnect($Server, 445, $null, $null).AsyncWaitHandle.WaitOne(40, $false) -or [system.Net.Sockets.TcpClient]::new().BeginConnect($Server, 445, $null, $null).AsyncWaitHandle.WaitOne(80, $false)

        if($Check -eq $true)
        {
            #$Server
            $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)
            $RegKey = $Reg.OpenSubKey("SYSTEM\\ControlSet001\\Control\\SystemInformation")
            $Value = $RegKey.GetValue("SystemManufacturer")

            if(($NULL -ne $Value) -AND ($Value -notlike "*VMware*"))
            {
                $Server
            }

            if($NULL -eq $Value)
            {
                $CheckModel = (Get-CimInstance -ComputerName $Server win32_computersystem -ErrorAction Stop).Model

                if(($NULL -ne $CheckModel) -AND ($CheckModel -notlike "*vmware*"))
                {
                    $Server
                }
            }
        }
    }
    Catch 
    {
    
    }

} | Wait-RSJob -ShowProgress -Timeout 1 | Receive-RSJob

#End Timestamp
$End = (Get-Date)

($End - $Start)