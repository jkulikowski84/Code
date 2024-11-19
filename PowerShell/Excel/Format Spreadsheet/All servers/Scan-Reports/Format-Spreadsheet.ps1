CLS

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

#------------------------------------ Remove Old Filtered Spreadsheet

$TEST_MAN_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\test\TEST_MAN_SERVER_GROUPS.xlsx"

if((Test-Path $TEST_MAN_SERVER_GROUPS) -eq $True)
{
    Remove-Item $TEST_MAN_SERVER_GROUPS
}

$PROD_MAN_TUE_10AM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\prod\TUE\PROD_MAN_TUE_10AM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_TUE_10AM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_TUE_10AM_SERVER_GROUPS
}

$PROD_MAN_TUE_6PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\prod\TUE\PROD_MAN_TUE_6PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_TUE_6PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_TUE_6PM_SERVER_GROUPS
}

$PROD_MAN_TUE_9PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\prod\TUE\PROD_MAN_TUE_9PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_TUE_9PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_TUE_9PM_SERVER_GROUPS
}

$PROD_MAN_THR_10AM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\prod\THR\PROD_MAN_THR_10AM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_THR_10AM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_THR_10AM_SERVER_GROUPS
}

$PROD_MAN_THR_6PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\prod\THR\PROD_MAN_THR_6PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_THR_6PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_THR_6PM_SERVER_GROUPS
}

$PROD_MAN_THR_9PM_SERVER_GROUPS = (Split-Path $script:MyInvocation.MyCommand.Path) + "\prod\THR\PROD_MAN_THR_9PM_SERVER_GROUPS.xlsx"

if((Test-Path $PROD_MAN_THR_9PM_SERVER_GROUPS) -eq $True)
{
    Remove-Item $PROD_MAN_THR_9PM_SERVER_GROUPS
}

$NewFilteredWorkbook = (Split-Path $script:MyInvocation.MyCommand.Path) + "\Filtered-Spreadsheet.xlsx"

if((Test-Path $NewFilteredWorkbook) -eq $True)
{
    Remove-Item $NewFilteredWorkbook
}

#------------------------------------  Setup Excel Variables

#The file we will be reading from
$ExcelFile = (Get-ChildItem -Path "$Path\*.xlsx").FullName

#Worksheet we are working on (by default this is the 1st tab)
$worksheet = (((New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList (New-Object -TypeName System.IO.FileStream -ArgumentList $ExcelFile,'Open','Read','ReadWrite')).Workbook).Worksheets[0]).Name

$ExcelServers = Import-Excel -Path $ExcelFile -WorkSheetname $worksheet -StartRow 1 -AsDate "Most recent discovery","Last Reboot Date"

#------------------------------------ Remove Duplicate entries and sort by Name

$SortedExcelServersList = ($ExcelServers | Sort-Object -Property Child -Unique)

#------------------------------------ Seperate Servers from DMZ Servers

$FilteredServers = ForEach($SortedExcelServerList in $SortedExcelServersList) {

    if(($($SortedExcelServerList.DMZ) -eq $true) -AND ($($SortedExcelServerList.child) -notlike "*.dmz.com"))
    {
        $SortedExcelServerList.child = [System.String]::Concat("$($SortedExcelServerList.child)",".dmz.com")
    }

    $SortedExcelServerList
}

#------------------------------------ Grab all servers from AD so we can use to compare against our list - also trimany whitespaces from output

$Servers = ((dsquery * -filter "(&(objectClass=Computer)(objectCategory=Computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(operatingSystem=*Server*))" -limit 0) | %{ if ($_ -match '^"CN=(.+?),\s*\w{1,2}=') { $matches[1] } }) | sort -Unique

#------------------------------------ Compare our list to servers in AD and filter out appliances

$FilteredServersResult = $Null

$FilteredServersResult = ForEach ($Item in $FilteredServers) 
{
    If (($item.child -in $Servers) -or ($item.DMZ -eq $True))
    {
        $Item
    }
}

#Create our new formatted Spreadsheet

$FilteredServersResult | Export-Excel -Path "Filtered-Spreadsheet.xlsx" -WorksheetName "All Servers" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

#------------------------------------ Create our grouped tabs

################ TEST

##Corepoint
$Test_Corepoint = ForEach($CorePointServer in $FilteredServersResult) {

    if((("$CorePointServer.Patch Window") -like "*TESTServerListManualReboot*") -And (("$CorePointServer.Parent") -like "*Corepoint*"))
    {
        $CorePointServer
    }
}

##ECW
$Test_ECW = ForEach($ECWtServer in $FilteredServersResult) {

    if((("$ECWtServer.Patch Window") -like "*TESTServerListManualReboot*") -And (("$ECWtServer.Parent") -like "*eClinicalWorks*"))
    {
        $ECWtServer
    }
}

##Epic
$Test_Epic = ForEach($EPicServer in $FilteredServersResult) {

    if((("$EPicServer.Patch Window") -like "*TESTServerListManualReboot*") -And (("$EPicServer.Parent") -like "*Epic*"))
    {
        $EPicServer
    }
}

##Lab
$Test_Lab = ForEach($LabServer in $FilteredServersResult) {

    if((("$LabServer.Patch Window") -like "*TESTServerListManualReboot*") -And ((("$LabServer.Parent") -like "*Data Innovations*") -OR (("$LabServer.Parent") -like "*Novanet*") -OR  (("$LabServer.Parent") -like "*QML*"))) 
    {
        $LabServer
    }
}

##Provisioning Worksations
$Test_Prov = ForEach($ProvServer in $FilteredServersResult) {

    if((("$ProvServer.Patch Window") -like "*TESTServerListManualReboot*") -And (("$ProvServer.Parent") -like "*AD Manager*"))
    {
        $ProvServer
    }
}

##Midas
$Test_Mids = ForEach($MidsServer in $FilteredServersResult) {

    if((("$MidsServer.Patch Window") -like "*TESTServerListManualReboot*") -And (("$MidsServer.Parent") -like "*Midas*"))
    {
        $MidsServer
    }
}

$Test_Corepoint | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "CorePoint" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_ECW | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "ECW" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Epic | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "EPIC" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Lab | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "LAB" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Prov | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "PROV Workstations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Test_Mids | Export-Excel -Path "$TEST_MAN_SERVER_GROUPS" -WorksheetName "Midas" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

################ Prod Tuesday

##10AM

##PROD 

##Automated IT Jobs
$Prod_TUE_10AM_AutomatedITJobs = ForEach($AutomatedITJobsServer in $FilteredServersResult) {

    if((("$AutomatedITJobsServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$AutomatedITJobsServer.Parent") -like "*Automated IT Jobs*"))
    {
        $AutomatedITJobsServer
    }
}

##Capsule
$Prod_TUE_10AM_Capsule = ForEach($CapsuleServer in $FilteredServersResult) {

    if((("$CapsuleServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$CapsuleServer.Parent") -like "*Capsule*"))
    {
        $CapsuleServer
    }
}

##DoseEdge
$Prod_TUE_10AM_DoseEdge = ForEach($DoseEdgeServer in $FilteredServersResult) {

    if((("$DoseEdgeServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$DoseEdgeServer.Parent") -like "*DoseEdge*"))
    {
        $DoseEdgeServer
    }
}

##Embla/Rembrandt
$Prod_TUE_10AM_Embla_Rembrandt = ForEach($Embla_RembrandtServer in $FilteredServersResult) {

    if((("$Embla_RembrandtServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$Embla_RembrandtServer.Parent") -like "*Embla*") -OR (("$Embla_RembrandtServer.Parent") -like "*Rembrandt*"))
    {
        $Embla_RembrandtServer
    }
}

##LAN/CoreInfrastructure
$Prod_TUE_10AM_LAN_CoreInfrastructure = ForEach($LAN_CoreInfrastructureServer in $FilteredServersResult) {

    if((("$LAN_CoreInfrastructureServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$LAN_CoreInfrastructureServer.Parent") -like "*LAN/Core Infrastructure*"))
    {
        $LAN_CoreInfrastructureServer
    }
}

##NurseCall
$Prod_TUE_10AM_NurseCall = ForEach($NurseCallServer in $FilteredServersResult) {

    if((("$NurseCallServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$NurseCallServer.Parent") -like "*Nurse Call*"))
    {
        $NurseCallServer
    }
}

##SleepWorks
$Prod_TUE_10AM_SleepWorks = ForEach($SleepWorksServer in $FilteredServersResult) {

    if((("$SleepWorksServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$SleepWorksServer.Parent") -like "*SleepWorks*"))
    {
        $SleepWorksServer
    }
}

#####LABS

##BioMearieux
$Prod_TUE_10AM_BioMearieux = ForEach($BioMearieuxServer in $FilteredServersResult) {

    if((("$BioMearieuxServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$BioMearieuxServer.Parent") -like "*Biomerieux*"))
    {
        $BioMearieuxServer
    }
}

##BioRad
$Prod_TUE_10AM_BioRad = ForEach($BioRadServer in $FilteredServersResult) {

    if((("$BioRadServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$BioRadServer.Parent") -like "*BioRad*"))
    {
        $BioRadServer
    }
}

##DataInnovations
$Prod_TUE_10AM_DataInnovations = ForEach($DataInnovationsServer in $FilteredServersResult) {

    if((("$DataInnovationsServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$DataInnovationsServer.Parent") -like "*Data Innovations*"))
    {
        $DataInnovationsServer
    }
}

##Novanet
$Prod_TUE_10AM_Novanet = ForEach($NovanetServer in $FilteredServersResult) {

    if((("$NovanetServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$NovanetServer.Parent") -like "*Novanet*"))
    {
        $NovanetServer
    }
}

##QML
$Prod_TUE_10AM_QML = ForEach($QMLServer in $FilteredServersResult) {

    if((("$QMLServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$QMLServer.Parent") -like "*QML*"))
    {
        $QMLServer
    }
}

##SCCSoftlab
$Prod_TUE_10AM_SCCSoftlab = ForEach($SCCSoftlabServer in $FilteredServersResult) {

    if((("$SCCSoftlabServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$SCCSoftlabServer.Parent") -like "*SCC Softlab*"))
    {
        $SCCSoftlabServer
    }
}

##TEG Manager
$Prod_TUE_10AM_TEGManager = ForEach($TEGManagerServer in $FilteredServersResult) {

    if((("$TEGManagerServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$TEGManagerServer.Parent") -like "*TEG Manager*"))
    {
        $TEGManagerServer
    }
}

##VoiceBrook
$Prod_TUE_10AM_VoiceBrook = ForEach($VoiceBrookServer in $FilteredServersResult) {

    if((("$VoiceBrookServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$VoiceBrookServer.Parent") -like "*Voicebrook*"))
    {
        $VoiceBrookServer
    }
}

##WAM
$Prod_TUE_10AM_WAM = ForEach($WAMServer in $FilteredServersResult) {

    if((("$WAMServer.Patch Window") -like "*TuesdayManualReboot - 10 am*") -And (("$WAMServer.Parent") -like "*WAM*"))
    {
        $WAMServer
    }
}

$Prod_TUE_10AM_AutomatedITJobs | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "AutomatedITJobs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Capsule | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "Capsule" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_DoseEdge | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "DoseEdge" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Embla_Rembrandt | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "Embla-Rembrandt" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_LAN_CoreInfrastructure | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "LAN-CoreInfrastructure" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_NurseCall | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "NurseCall" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_SleepWorks | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "SleepWorks" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_BioMearieux | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "BioMearieux" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_BioRad | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "BioRad" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_DataInnovations | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "DataInnovations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Novanet | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "Novanet" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_QML | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "QML" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_SCCSoftlab | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "SCCSoftlab" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_TEGManager | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "TEGManager" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_VoiceBrook | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "VoiceBrook" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_WAM | Export-Excel -Path "$PROD_MAN_TUE_10AM_SERVER_GROUPS" -WorksheetName "WAM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

##6PM

##Capsule
$Prod_TUE_6PM_Capsule = ForEach($CapsuleServer in $FilteredServersResult) {

    if((("$CapsuleServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*") -And (("$CapsuleServer.Parent") -like "*Capsule*"))
    {
        $CapsuleServer
    }
}

##Corepoint
$Prod_TUE_6PM_Corepoint = ForEach($CorepointServer in $FilteredServersResult) {

    if((("$CorepointServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*") -And (("$CorepointServer.Parent") -like "*Corepoint*"))
    {
        $CorepointServer
    }
}

##Exchange
$Prod_TUE_6PM_Exchange = ForEach($ExchangeServer in $FilteredServersResult) {

    if((("$ExchangeServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*") -And (("$ExchangeServer.Parent") -like "*Exchange*"))
    {
        $ExchangeServer
    }
}

##Midas
$Prod_TUE_6PM_Midas = ForEach($MidasServer in $FilteredServersResult) {

    if((("$MidasServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*") -And (("$MidasServer.Parent") -like "*Midas*"))
    {
        $MidasServer
    }
}

##Spacelabs
$Prod_TUE_6PM_Spacelabs = ForEach($SpacelabsServer in $FilteredServersResult) {

    if((("$SpacelabsServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*") -And (("$SpacelabsServer.Parent") -like "*Spacelabs*"))
    {
        $SpacelabsServer
    }
}

##Stryker
$Prod_TUE_6PM_Stryker = ForEach($StrykerServer in $FilteredServersResult) {

    if((("$StrykerServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*") -And (("$StrykerServer.Parent") -like "*Stryker*"))
    {
        $StrykerServer
    }
}

##SynaptiveMedical_ImageDrive
$Prod_TUE_6PM_SynaptiveMedical_ImageDrive = ForEach($SynaptiveMedical_ImageDriveServer in $FilteredServersResult) {

    if((("$SynaptiveMedical_ImageDriveServer.Patch Window") -like "*TuesdayManualReboot - 6 pm*") -And (("$SynaptiveMedical_ImageDriveServer.Parent") -like "*Synaptive Medical*") -OR (("$SynaptiveMedical_ImageDriveServer.Parent") -like "*Image Drive*"))
    {
        $SynaptiveMedical_ImageDriveServer
    }
}

$Prod_TUE_6PM_Capsule | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Capsule" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Corepoint | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Corepoint" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Exchange | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Exchange" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Midas | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Midas" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Spacelabs | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Spacelabs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Stryker | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "Stryker" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_SynaptiveMedical_ImageDrive | Export-Excel -Path "$PROD_MAN_TUE_6PM_SERVER_GROUPS" -WorksheetName "SynaptiveMedical-ImageDrive" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

##9PM

##AtlasQA
$Prod_TUE_9PM_AtlasQA = ForEach($AtlasQAServer in $FilteredServersResult) {

    if((("$AtlasQAServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$AtlasQAServer.Parent") -like "*Atlas*"))
    {
        $AtlasQAServer
    }
}

##Avaya
$Prod_TUE_9PM_Avaya = ForEach($AvayaServer in $FilteredServersResult) {

    if((("$AvayaServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$AvayaServer.Parent") -like "*AVST*"))
    {
        $AvayaServer
    }
}

##CBORD_Aramark
$Prod_TUE_9PM_CBORD_Aramark = ForEach($CBORD_AramarkServer in $FilteredServersResult) {

    if((("$CBORD_AramarkServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$CBORD_AramarkServer.Parent") -like "*CBORD*") -OR (("$CBORD_AramarkServer.Parent") -like "*Aramark*"))
    {
        $CBORD_AramarkServer
    }
}

##C-Rad
$Prod_TUE_9PM_CRad = ForEach($CRadServer in $FilteredServersResult) {

    if((("$CRadServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$CRadServer.Parent") -like "*C-RAD*") )
    {
        $CRadServer
    }
}

##MIM
$Prod_TUE_9PM_MIM = ForEach($MIMServer in $FilteredServersResult) {

    if((("$MIMServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$MIMServer.Parent") -like "*MIM*") )
    {
        $MIMServer
    }
}

##Nuance
$Prod_TUE_9PM_Nuance = ForEach($NuanceServer in $FilteredServersResult) {

    if((("$NuanceServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$NuanceServer.Parent") -like "*Nuance*") )
    {
        $NuanceServer
    }
}

##Obix
$Prod_TUE_9PM_Obix = ForEach($ObixServer in $FilteredServersResult) {

    if((("$ObixServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$ObixServer.Parent") -like "*Obix*") )
    {
        $ObixServer
    }
}

##PaceArt
$Prod_TUE_9PM_PaceArt = ForEach($PaceArtServer in $FilteredServersResult) {

    if((("$PaceArtServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$PaceArtServer.Parent") -like "*PaceArt*") )
    {
        $PaceArtServer
    }
}

##RayStation
$Prod_TUE_9PM_RayStation = ForEach($RayStationServer in $FilteredServersResult) {

    if((("$RayStationServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$RayStationServer.Parent") -like "*Ray Station*") )
    {
        $RayStationServer
    }
}

##Syngo
$Prod_TUE_9PM_Syngo = ForEach($SyngoServer in $FilteredServersResult) {

    if((("$SyngoServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$SyngoServer.Parent") -like "*Syngo*") )
    {
        $SyngoServer
    }
}

##SunCheck
$Prod_TUE_9PM_SunCheck = ForEach($SunCheckServer in $FilteredServersResult) {

    if((("$SunCheckServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$SunCheckServer.Parent") -like "*SunCheck*") )
    {
        $SunCheckServer
    }
}

##Vitrea
$Prod_TUE_9PM_Vitrea = ForEach($VitreaServer in $FilteredServersResult) {

    if((("$VitreaServer.Patch Window") -like "*TuesdayManualReboot - 9 pm*") -And (("$VitreaServer.Parent") -like "*Vitrea*") )
    {
        $VitreaServer
    }
}

$Prod_TUE_9PM_AtlasQA | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "AtlasQA" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Avaya | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Avaya" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_CBORD_Aramark | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "CBORD-Aramark" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_CRad | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "C-RAD" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_MIM | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "MIM" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Nuance | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Nuance" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Obix | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Obix" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_PaceArt | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "PaceArt" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_RayStation | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "RayStation" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Syngo | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Syngo" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_SunCheck | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "SunCheck" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Vitrea | Export-Excel -Path "$PROD_MAN_TUE_9PM_SERVER_GROUPS" -WorksheetName "Vitrea" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"


################ Prod Thursday

##10AM

##Kronos
$Prod_TUE_10AM_Kronos = ForEach($KronosServer in $FilteredServersResult) {

    if((("$KronosServer.Patch Window") -like "*ThursdayManualReboot - 10 am*") -And (("$KronosServer.Parent") -like "*Kronos*") )
    {
        $KronosServer
    }
}

##Spok
$Prod_TUE_10AM_Spok = ForEach($SpokServer in $FilteredServersResult) {

    if((("$SpokServer.Patch Window") -like "*ThursdayManualReboot - 10 am*") -And (("$SpokServer.Parent") -like "*Spok*") )
    {
        $SpokServer
    }
}

$Prod_TUE_10AM_Kronos | Export-Excel -Path "$PROD_MAN_THR_10AM_SERVER_GROUPS" -WorksheetName "Kronos" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_10AM_Spok | Export-Excel -Path "$PROD_MAN_THR_10AM_SERVER_GROUPS" -WorksheetName "Spok" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

##6PM

##Axis-Pats
$Prod_TUE_6PM_Axis_Pats = ForEach($Axis_PatsServer in $FilteredServersResult) {

    if((("$Axis_PatsServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*") -And (("$Axis_PatsServer.Parent") -like "*Axis*") )
    {
        $Axis_PatsServer
    }
}

##Elekta Mosaiq
$Prod_TUE_6PM_Elekta_Mosaiq = ForEach($Elekta_MosaiqServer in $FilteredServersResult) {

    if((("$Elekta_MosaiqServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*") -And (("$Elekta_MosaiqServer.Parent") -like "*Elekta Mosaiq*") )
    {
        $Elekta_MosaiqServer
    }
}

##Epic
$Prod_TUE_6PM_Epic = ForEach($EpicServer in $FilteredServersResult) {

    if((("$EpicServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*") -And (("$EpicServer.Parent") -like "*Epic*") )
    {
        $EpicServer
    }
}

##Exchange
$Prod_TUE_6PM_Exchange = ForEach($ExchangeServer in $FilteredServersResult) {

    if((("$ExchangeServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*") -And (("$ExchangeServer.Parent") -like "*Exchange*") )
    {
        $ExchangeServer
    }
}

##GE PACS
$Prod_TUE_6PM_GEPACS = ForEach($GEPACSServer in $FilteredServersResult) {

    if((("$GEPACSServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*") -And (("$GEPACSServer.Parent") -like "*GE PACS*") )
    {
        $GEPACSServer
    }
}

##RCA
$Prod_TUE_6PM_RCA = ForEach($RCAServer in $FilteredServersResult) {

    if((("$RCAServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*") -And (("$RCAServer.Parent") -like "*RCA*") )
    {
        $RCAServer
    }
}

##Varian
$Prod_TUE_6PM_Varian = ForEach($VarianServer in $FilteredServersResult) {

    if((("$VarianServer.Patch Window") -like "*ThursdayManualReboot - 6 pm*") -And (("$VarianServer.Parent") -like "*Varian*") )
    {
        $VarianServer
    }
}

$Prod_TUE_6PM_Axis_Pats | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Axis_Pats" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Elekta_Mosaiq | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Elekta_Mosaiq" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Epic | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "EPIC" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Exchange | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Exchange" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_GEPACS | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "GE-Pacs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_RCA | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "RCA" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_6PM_Varian | Export-Excel -Path "$PROD_MAN_THR_6PM_SERVER_GROUPS" -WorksheetName "Varian" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

##9PM

##3M
$Prod_TUE_9PM_3M = ForEach($3MServer in $FilteredServersResult) {

    if((("$3MServer.Patch Window") -like "*ThursdayManualReboot - 9 pm*") -And (("$3MServer.Parent") -like "*3M*") )
    {
        $3MServer
    }
}

##Dexa
$Prod_TUE_9PM_Dexa = ForEach($DexaServer in $FilteredServersResult) {

    if((("$DexaServer.Patch Window") -like "*ThursdayManualReboot - 9 pm*") -And (("$DexaServer.Parent") -like "*Dexa*") )
    {
        $DexaServer
    }
}

##GE Pacs
$Prod_TUE_9PM_GE_Pacs = ForEach($GE_PacsServer in $FilteredServersResult) {

    if((("$GE_PacsServer.Patch Window") -like "*ThursdayManualReboot - 9 pm*") -And (("$GE_PacsServer.Parent") -like "*GE Pacs*") )
    {
        $GE_PacsServer
    }
}

##OneContent/ROI
$Prod_TUE_9PM_OneContent_ROI = ForEach($OneContent_ROIServer in $FilteredServersResult) {

    if((("$OneContent_ROIServer.Patch Window") -like "*ThursdayManualReboot - 9 pm*") -And ((("$OneContent_ROIServer.Parent") -like "*OneContent*") -OR (("$OneContent_ROIServer.Parent") -like "*ROI*") ))
    {
        $OneContent_ROIServer
    }
}

##Quick Charge
$Prod_TUE_9PM_Quick_Charge = ForEach($Quick_ChargeServer in $FilteredServersResult) {

    if((("$Quick_ChargeServer.Patch Window") -like "*ThursdayManualReboot - 9 pm*") -And (("$Quick_ChargeServer.Parent") -like "*Quick Charge*"))
    {
        $Quick_ChargeServer
    }
}

$Prod_TUE_9PM_3M | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "3M" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Dexa | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "Dexa" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_GE_Pacs | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "GE-Pacs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_OneContent_ROI | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "OneContent-ROI" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"
$Prod_TUE_9PM_Quick_Charge | Export-Excel -Path "$PROD_MAN_THR_9PM_SERVER_GROUPS" -WorksheetName "Quick-Charge" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle "Light1"

$End =  (Get-Date)

$End - $Start
