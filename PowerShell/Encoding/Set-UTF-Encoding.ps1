CLS

$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
$source = "C:\GitHub\Code"
$destination = "C:\GitHub\Test"

if (!(Test-Path $destination))
{
    write-host "(Note: target folder created!) "
    new-item -type directory -path $destination -Force | Out-Null
}

foreach ($i in Get-ChildItem $Source -Recurse -Force) 
{
    if ($i.PSIsContainer) 
    {
        continue
    }
    
    $path = $i.DirectoryName.Replace($source, $destination)
    $name = $i.Fullname.Replace($source, $destination)

    if ( !(Test-Path $path) ) 
    {
        New-Item -Path $path -ItemType directory
    }

    $content = get-content $i.Fullname

    if ( $content -ne $null ) 
    {

        [System.IO.File]::WriteAllLines($name, $content, $Utf8NoBomEncoding)
    } 
    else 
    {
        Write-Host "No content from: $i"   
    }
}
