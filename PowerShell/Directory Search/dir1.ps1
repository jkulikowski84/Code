CLS

#$directory = "\\nch_home\IT\Server Team\WindowsUpdates"
$directory = "\\pwvs01ldps0001\Fileserver\Software"

#3.7623048
Measure-Command {
    cmd.exe /c dir "$directory" /a-d /-c /s
}

#4.3868718
Measure-Command {
    cmd /c dir "$directory" /B /S /A-D
}

#6.6459505
Measure-Command {
    robocopy $directory NULL /L /MIR /NJH /FP /NC /NDL /NS /NJS
}

#5.7641184
Measure-Command {
    robocopy $directory NULL /L /S /NJH /BYTES /FP /NC /NDL /XJ /TS /R:0 /W:0
}

#5.1776761
Measure-Command {
    (dir -Path $directory -Recurse).FullName
}

#3.002
measure-command{
    [System.IO.Directory]::EnumerateFiles("$directory","*.*","AllDirectories")
}

#3.3726
measure-command{
    [System.IO.Directory]::EnumerateFileSystemEntries("$directory", "*.*", "AllDirectories")
}

#3.072
measure-command{
    (Get-Item $directory).EnumerateFiles("*.*", 'AllDirectories')
}

