CLS

Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class Win {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@

[string]$vpncliAbsolutePath = 'C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe'
[string]$vpnuiAbsolutePath = 'C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe'

#Close out of All VPN processes (to avoid conflict)

Get-Process vpn* -ErrorAction SilentlyContinue | ForEach-Object { 

    if(($_.ProcessName -eq "vpnui") -OR ($_.ProcessName -eq "vpncli"))
    {
        Try
        {
            Stop-Process $_.Id -Force
        }
        Catch
        {
            taskkill /IM $_.ProcessName /f | Out-Null 
        }
    }
}

#Pause 2 seconds
Sleep -Seconds 2

Start-Process -WindowStyle Normal -FilePath $vpncliAbsolutePath -ArgumentList "connect SECURE-VPN"

#Grab the Process Handle
do
{
    $h = (Get-Process vpncli).MainWindowHandle
}while(($NULL -eq $h) -OR ($h -eq 0))

#Set the Window to the front
[Win]::SetForegroundWindow($h) | Out-Null

#=========== Connect to the VPN

#Press Enter to select our username
[System.Windows.Forms.SendKeys]::SendWait("{Enter}")

#Input the password
$PW = (Import-Clixml "C:\Scripts\VPN-Connection\creds.xml").GetNetworkCredential().password
[System.Windows.Forms.SendKeys]::SendWait("$PW{Enter}")

#Pause 2 seconds
Sleep -Seconds 2

#Launch our VPN UI
Start-Process -WindowStyle Minimized -FilePath $vpnuiAbsolutePath