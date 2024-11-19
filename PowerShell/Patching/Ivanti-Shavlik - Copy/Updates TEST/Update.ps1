CLS

#------------- MSCatalog (For Querying Windows Catalog)

if((Get-Module -ListAvailable -Name "MSCatalog") -or (Get-Module -Name "MSCatalog"))
{
        Import-Module MSCatalog
}
else
{	
    Install-Module -Name MSCatalog -Scope CurrentUser -Force -Confirm:$False
	Import-Module MSCatalog
}

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
$ResultsFile = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Results.csv"
$DMZAccount = (whoami).replace("domain","dmz.com")

#------------------------------------  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

$ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1

#------------------------------------ Populate our variable with data from spreadsheet

$ExcelServersList = foreach($ExcelServer in $ExcelServers) {
    $ExcelServer | Select-Object @{Name="ServerName";Expression={$_.Child}}, "Primary", @{Name="PatchWindow";Expression={$_."Patch Window"}}, @{Name="TestServer";Expression={$_."Test Server"}}, "DMZ", @{Name="OS";Expression={((($_."Operating System").Replace("Windows ","")) -replace "(\s\S+)$") }}
}

#------------------------------------ Remove Duplicate entries

$SortedExcelServersList = ($ExcelServersList | Sort-Object -Property ServerName -Unique)

#------------------------------------ Seperate Servers from DMZ Servers

$FilteredServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if(($($SortedExcelServerList.DMZ) -eq $true) -AND ($($SortedExcelServerList.ServerName) -notlike "*.dmz.com"))
    {
        $SortedExcelServerList.ServerName = [System.String]::Concat("$($SortedExcelServerList.ServerName)",".dmz.com")
    }

    $SortedExcelServerList
}

#------------------------------------ Grab all servers from AD so we can use to compare against our list - also trimany whitespaces from output

$Servers = (dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0 -attr Name | sort).trim()

#------------------------------------ Compare our list to servers in AD and filter out appliances

$FilteredServersResult = ForEach($Item in $FilteredServers) {

    If (($item.ServerName -in $Servers) -or ($item.DMZ -eq $True))
    {
        $Item
    }
}

#---------------------------- Perform our search. In this case all Monthly Updates and filter it to only show updates for Servers

$Search = (Get-Date -Format "yyyy-MM")

$Updates = (Get-MSCatalogUpdate -Search $Search -AllPages -ErrorAction Stop)

if($NULL -eq $Updates)
{
	#If NULL, exit because we have no new updates
	exit
}

$WinCatalog = foreach($Update in $Updates) {
    $Update | Where { (($_.Title -like "*Server*") -OR ($_.Products -like "*Server*")) -AND (($_.Classification -eq "Critical Updates") -OR ($_.Classification -eq "Security Updates")) } | Select-Object @{Name="Title";Expression={([regex]::Matches(($($_).Title), '(?<=\().+?(?=\))')).Value}}, @{Name="Products";Expression={(($_."Products").Replace("Windows Server ",""))}}, "Classification", "LastUpdated", "Size"
}

#------------------------------------ Multithreading Magic

$DMZAccount = Get-Credential ((whoami).replace("domain","dmz.com"))

$FilteredServersResult | Start-RSJob -Throttle $($FilteredServersResult.Count) -Batch "Test" -ErrorAction Stop -ScriptBlock {
    
    Param($Server)
    $Check = $False

    #------------------------------------ Ping servers to make sure they're responsive

    Try
    {
        if($NULL -ne (Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$($Server.servername)' AND Timeout=100" -ErrorAction Stop).ResponseTime)
        {
            #------------------------------------ Get Server KB's

            $NeedsUpdates = $ServerKBs = $NULL

            Try
            {
                [ScriptBlock]$SB = {
                    
                    $DateSearch = (Get-Date -Format "*MM*yyyy*")

                    $updateSession = New-Object -ComObject Microsoft.Update.Session
                    $updateSearch = $updateSession.CreateUpdateSearcher()
                    $updateCount = $updateSearch.GetTotalHistoryCount()

                    $ServerUpdates = ($updateSearch.QueryHistory(0,$updateCount) | Select Date,Title,Description | Where-Object { ($PSItem.Title) -AND ($_.Date -like $DateSearch) })

                    ([regex]::Matches(($($ServerUpdates).Title), '(?<=\().+?(?=\))')).Value
                }

                if($Server.DMZ -eq $TRUE)
                {
                    $ServerKBs = Invoke-Command -ComputerName $($Server.servername) -Credential $using:DMZAccount -ErrorAction Stop -ScriptBlock $SB | where {$_ -like "KB*"}
                    $Check = $True
                }
                else
                {
                    $ServerKBs = Invoke-Command -ComputerName $($Server.servername) -ErrorAction Stop -ScriptBlock $SB | where {$_ -like "KB*"}
                    $Check = $True
                }
            }
            Catch
            {
                if($Check -eq $False)
                {
                    Try
                    {
                        $DateSearch = (Get-Date -Format "*MM*yyyy*")

                        $ServerKBs = (((Get-HotFix -ComputerName $($Server.servername) -ErrorAction Stop) | Select-Object HotFixID,InstalledOn) | Where {($_.InstalledOn -like "$DateSearch" )}).HotFixID
                        $Check = $True

                    }
                    Catch
                    {
                        ($Server | Add-Member -NotePropertyMembers @{"Error" = [string]$Error} -Force -PassThru) | Export-Csv -Path $using:ErrorFile -NoTypeInformation -Force -Append
                    }
                }
            }

            #---------------------------- Compare Updates on server to WIndows Update Catalog to determine which updates are missing

            $NeedsUpdates = foreach($item in $using:WinCatalog) {

                #Match up the Update for the OS of the server

                if($item.Products -eq $($Server.OS))
                {
                    #Now check if the update is missing
                    if(($NULL -eq $ServerKBs) -OR ($item.Title -notin $ServerKBs))
                    {
                        $item.Title
                    }
                }
            }

            if($NULL -ne $NeedsUpdates)
            {
                ($Server | Add-Member -NotePropertyMembers @{"KBs" = (@($($NeedsUpdates)) -join ', ')} -PassThru)
            }
        }
    }
    Catch
    {
        ($Server | Add-Member -NotePropertyMembers @{"Error" = [string]$Error} -Force -PassThru) | Export-Csv -Path $using:ErrorFile -NoTypeInformation -Force -Append
    }

} | Wait-RSJob -ShowProgress -Timeout 30 | Receive-RSJob | Export-Csv -Path $ResultsFile -NoTypeInformation -Force

$End =  (Get-Date)

$End - $Start
