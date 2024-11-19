CLS

#------------- PoshRSJob (multitasking)

if((Get-Module -ListAvailable -Name "PoshRSJob") -or (Get-Module -Name "PoshRSJob"))
{
        Import-Module PoshRSJob
}
else
{	
    Install-Module -Name PoshRSJob -Scope CurrentUser -Force -Confirm:$False
	Import-Module PoshRSJob
}


#------------- ImportExcel Module 

if((Get-Module -ListAvailable -Name "ImportExcel") -or (Get-Module -Name "ImportExcel"))
{
        Import-Module ImportExcel
}
else
{
    #Install NuGet (Prerequisite) first
	Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$False
	
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$False
	Import-Module ImportExcel
}

#Clear screen again
CLS

#----------------------------------------------------------------------------------------------------------------

#Start Timestamp
$Start = Get-Date

#Global Variables
$Path = (Split-Path $script:MyInvocation.MyCommand.Path)
$ErrorFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\ERROR.csv"

$Results = "$Path\Results.xlsx"

if((Test-Path $Results) -eq $True)
{
    Remove-Item $Results
}

if((Test-Path $ErrorFile) -eq $True)
{
    Remove-Item $ErrorFile
}

#------------------------------------  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

$ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

#------------------------------------ Populate our variable with data from spreadsheet

$ExcelServersList = foreach($ExcelServer in $ExcelServers) {
    $ExcelServer | Select-Object @{Name="ServerName";Expression={$_.Child}}, "Primary", @{Name="PatchWindow";Expression={$_."Patch Window"}}, @{Name="TestServer";Expression={$_."Test Server"}}, "DMZ", "Operating System"
}

#------------------------------------ Remove Duplicate entries

$SortedExcelServersList = ($ExcelServersList | Sort-Object -Property ServerName -Unique)

#------------------------------------ Seperate Servers from DMZ Servers

$FilteredServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if($($SortedExcelServerList.DMZ) -eq $true)
    {
        $SortedExcelServerList.ServerName = [System.String]::Concat("$($SortedExcelServerList.ServerName)",".dmz.com")
    }

    $SortedExcelServerList
}

#------------------------------------ Grab all servers from AD so we can use to compare against our list - also trim any whitespaces from output

$Servers = (dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0 -attr Name | sort).trim()

#------------------------------------ Compare our list to servers in AD and filter out appliances

$FilteredServersResult = $Null

$FilteredServersResult = ForEach ($Item in $FilteredServers) 
{
    If (($item.servername -in $Servers) -or ($item.DMZ -eq $True))
    {
        $Item
    }
}

#------------------------------------ Multithreading Magic

$FilteredServersResult | Start-RSJob -Throttle 25 -ScriptBlock {
    Param($Server)

    #Ping servers to make sure they're responsive
    if($NULL -ne (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$($Server.servername)' AND Timeout=100").ResponseTime)
    { 
        Try
        {
            #$ServerPath = "\\$($Server.ServerName)\c$\Windows\ServiceProfiles\LocalService\NTUSER.DAT"
            $ServerPath = "\\$($Server.ServerName)\c$\Windows\ServiceProfiles\LocalService"

            #If you're checking file date
            #$Check = ((Get-Date) - (dir $ServerPath -force -ErrorAction Stop).LastWriteTime).days
            $Check = ((Get-Date) - ((dir $ServerPath -force -ErrorAction Stop -Filter "NTUSER*") | sort LastWriteTime -Descending).LastWriteTime[0]).Days
        }
        Catch
        {
            Try
            {
                $Check = ((Get-Date) - (Get-WmiObject win32_operatingsystem -ComputerName $($Server.ServerName) -ErrorAction Stop | select csname, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}).LastBootUpTime).days
            }
            Catch
            {                
                if($Error.exception -like "*")
                {
                    ($Server | Add-Member -NotePropertyMembers @{"Error" = [string]$Error} -PassThru) | Export-Csv -Path $using:ErrorFile -NoTypeInformation -Force -Append
                }

                #$Error | Select –Property *
            }
        }
        
        #Check if server hasn't been rebooted in greater than or equal to 25 days.
        if($Check -ge 25)
        {
            $Server | Add-Member -NotePropertyMembers @{"Day's Online" = $Check} -PassThru
        }
    }
} | Wait-RSJob -ShowProgress -Timeout 30 | Receive-RSJob | Export-Excel -Path $Results -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
#} | Wait-RSJob -ShowProgress | Receive-RSJob | Export-Csv -Path "$Path\Results.csv" -NoTypeInformation -Force 

$End =  (Get-Date)

$End - $Start
