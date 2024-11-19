CLS

$IP = "123.123.123.123"
$statistic = $Null
$date = Get-Date
$URL = "https://${IP}"

$WebRequest = [Net.WebRequest]::Create($URL)
$WebRequest.UseDefaultCredentials = $true
$WebRequest.PreAuthenticate = $true

$AllArray = @()

Try
{
    $WebResponse = $WebRequest.GetResponse()
    $Cert = [Security.Cryptography.X509Certificates.X509Certificate2]$WebRequest.ServicePoint.Certificate.Handle
    $statistic = $cert.Subject
    $expiry = $cert.NotAfter
    $remaining = $expiry - $date
    $Statistic = $remaining.days
}
Catch
{
    # Write-Host "Web request failed" -ForegroundColor Red
    # Write-Host "Attempting to get cert info regardless..." -ForegroundColor Yellow

    $Cert = [Security.Cryptography.X509Certificates.X509Certificate2]$WebRequest.ServicePoint.Certificate.Handle
    $CN = $cert.Subject
    $expiry = $cert.NotAfter
    $remaining = $expiry - $date
    $Statistic = $remaining.Days

    If($statistic -lt "-2000")
    {
        Clear-Variable statistic
    }
}

If($Statistic -ne $null)
{
    $FormattedExpiry = $expiry.ToString("dd/MM/yyyy")
    $Message = "Certificate $CN will expire on $FormattedExpiry, $statistic days left"
    Write-Host "Statistic: $statistic"
    Write-Host "Message: $message"
    Exit 0;
}

Function Get-Direct
{
    If($statistic -eq $Null)
    {
        #Write-Host "Trying direct cert store script" -ForegroundColor Yellow
        $server = $url.Replace('https://','')
        $objStore = new-object System.Security.Cryptography.X509Certificates.X509Store("\\$Server\MY","LocalMachine")
        $objStore.open("ReadOnly")
        $Cert = $objStore.Certificates | sort notafter
        $CN = $Cert.subject[0]
        $Expiry = $Cert.NotAfter[0]
        $Remaining = $expiry - $date
        $statistic = $remaining.Days

        If($statistic -lt "-2000")
        {
            Clear-Variable statistic
        }

        If($statistic -eq $Null)
        {
            Write-Host "Statistic.ExitCode: 1"
            Exit 1;
        }
        Else
        {
            $FormattedExpiry = $expiry.ToString("dd/MM/yyyy")
            $Message = "Certificate $CN will expire on $FormattedExpiry, $statistic days left"
            Write-Host "Statistic: $statistic"
            Write-Host "Message: $message"
            Exit 0;
        }
    }
}

If($statistic -eq $null)
{
    Get-Direct
}
