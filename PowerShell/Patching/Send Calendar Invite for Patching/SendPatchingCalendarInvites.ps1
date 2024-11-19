CLS

$Path = (Split-Path $script:MyInvocation.MyCommand.Path)

$SmtpClient = new-object system.net.mail.smtpClient
$SmtpClient.Host = "smtp.domain.com"

$Mail = New-Object System.Net.Mail.MailMessage

$Mail.From = "PatchingInfo@domain.com"
$Mail.IsBodyHtml = $true

#====================== Files we will be attaching to our Calendar invite

if($NULL -eq $ValidationFile)
{
    $ValidationFile = "$Path\Patching Validation in Business Service.docx"
}
if($NULL -eq $ServerList)
{
    $ServerList = "$Path\Filtered-Spreadsheet.xlsx"
}

#====================== Calculate dates for the invite
    
#Get current Month and Year
$Month = (get-Date).month
$Year = (get-Date).year

#Get the first day of the month
$FirstDayofMonth = [datetime] ([string]$Month + "/1/" + [string]$Year)

#Grab the Wednesday after Windows Patch Tuesday (2nd Tuesay of the month) for Test Patching
$Wed = (0..30 | % {$firstdayofmonth.adddays($_) } | ? {$_.dayofweek -like "Tue*"})[1].AddDays(1)

#Check if the date for test Patching is correct.
Write-output "Test Patching is scheduled for $($Wed.ToString('MM/dd/yyyy'))"
$TestDay = Read-Host -Prompt "If the above date is correct, hit 'Enter' to continue, otherwise type in the number of the day test patching will take place on."

if([string]::IsNullOrEmpty($TestDay))
{
    #Do nothing, just continue
}
else
{
    $Wed = Get-Date -Date "$($year)-$($month)-$($TestDay)T00:00:00"
}

CLS

#Prod Patching *USUALLY* takes place on the 3rd week (after Test Patching) on Tuesday and Thursday
$Tue = ($Wed).AddDays(6)

#Check if the date for Tue Prod Patching is correct.
Write-output "Tue Prod Patching is scheduled for $($Tue.ToString('MM/dd/yyyy'))"
$ProdDay1 = Read-Host -Prompt "If the above date is correct, hit 'Enter' to continue, otherwise type in the number of the day test patching will take place on."

if([string]::IsNullOrEmpty($ProdDay1))
{
    #Do nothing, just continue
}
else
{
    $Tue = Get-Date -Date "$($year)-$($month)-$($ProdDay1)T00:00:00"
}

CLS

$Thr = ($Tue).AddDays(2)

#Check if the date for Thr Prod Patching is correct.
Write-output "Thr Prod Patching is scheduled for $($Thr.ToString('MM/dd/yyyy'))"
$ProdDay2 = Read-Host -Prompt "If the above date is correct, hit 'Enter' to continue, otherwise type in the number of the day test patching will take place on."

if([string]::IsNullOrEmpty($ProdDay2))
{
    #Do nothing, just continue
}
else
{
    $Thr = Get-Date -Date "$($year)-$($month)-$($ProdDay2)T00:00:00"
}

CLS

#-------------------
#$PatchWindows = @("Test10AM","Test6PM","ProdTUE10AM","ProdTUE6PM","ProdTUE9PM","ProdTHR10AM","ProdTHR6PM","ProdTHR9PM")
$PatchWindows = @()
$FutureDate = $False

#Skip old dates
if($Wed -ge $(Get-Date))
{
    Write-Output "Test Patching will take place on: $($Wed.ToString('MM/dd/yyyy'))"
    Write-Output ""
    $PatchWindows += @("Test10AM","Test6PM")
    $FutureDate = $True
}

if($Tue -ge $(Get-Date))
{
    Write-Output "Tue Prod Patching will take place on: $($Tue.ToString('MM/dd/yyyy'))"
    Write-Output ""
    $PatchWindows += @("ProdTUE10AM","ProdTUE6PM","ProdTUE9PM")
    $FutureDate = $True
}

if($Thr -ge $(Get-Date))
{
    Write-Output "Thr Prod Patching will take place on: $($Thr.ToString('MM/dd/yyyy'))"
    Write-Output ""
    $PatchWindows += @("ProdTHR10AM","ProdTHR6PM","ProdTHR9PM")
    $FutureDate = $True
}

if($FutureDate -eq $True)
{
    foreach($PatchWindow in $PatchWindows)
    {   
        $Attendees = $PatchStart = $vCalTitle = $StartTime = $EndTime = $NULL

        Switch ($PatchWindow)
        {
            "Test10AM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "10 AM"
                $vCalTitle = "Placeholder for Wednesday Morning Windows Test Servers Patching - Starting at 10 AM"
                [string]$StartTime = $(($Wed).AddHours(10)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Wed).AddHours(12)).ToString('yyyyMMddTHHmmss')
            }
            "Test6PM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "6 PM"
                $vCalTitle = "Placeholder for Wednesday Windows Test Servers Patching - Starting at 6 PM"
                [string]$StartTime = $(($Wed).AddHours(18)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Wed).AddHours(20)).ToString('yyyyMMddTHHmmss')
            }
            "ProdTUE10AM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "10 AM"
                $vCalTitle = "Placeholder for Tuesday Morning Windows Prod Servers Patching - Starting at 10 AM"
                [string]$StartTime = $(($Tue).AddHours(10)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Tue).AddHours(12)).ToString('yyyyMMddTHHmmss')
            }
            "ProdTUE6PM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "6 PM"
                $vCalTitle = "Placeholder for Tuesday Windows Prod Servers Patching - Starting at 6 PM"
                [string]$StartTime = $(($Tue).AddHours(18)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Tue).AddHours(20)).ToString('yyyyMMddTHHmmss')
            }
            "ProdTUE9PM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "9 PM"
                $vCalTitle = "Placeholder for Tuesday Windows Prod Servers Patching - Starting at 9 PM"
                [string]$StartTime = $(($Tue).AddHours(21)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Tue).AddHours(23)).ToString('yyyyMMddTHHmmss')
            }
            "ProdTHR10AM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "10 AM"
                $vCalTitle = "Placeholder for Thursday Morning Windows Prod Servers Patching - Starting at 10 AM"
                [string]$StartTime = $(($Thr).AddHours(10)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Thr).AddHours(12)).ToString('yyyyMMddTHHmmss')
            }
            "ProdTHR6PM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "6 PM"
                $vCalTitle = "Placeholder for Thursday Windows Prod Servers Patching - Starting at 6 PM"
                [string]$StartTime = $(($Thr).AddHours(18)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Thr).AddHours(20)).ToString('yyyyMMddTHHmmss')
            }
            "ProdTHR9PM"
            {
                $Attendees = @("ATTENDEE;CN='Person, Someone';RSVP=TRUE:mailto:someoneA@domain.com`n
                ATTENDEE;CN='Person, Someone';ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:someoneB@domain.com")
                $PatchStart = "9 PM"
                $vCalTitle = "Placeholder for Thursday Windows Prod Servers Patching - Starting at 9 PM"
                [string]$StartTime = $(($Thr).AddHours(21)).ToString('yyyyMMddTHHmmss')
                [string]$EndTime = $(($Thr).AddHours(23)).ToString('yyyyMMddTHHmmss')
            }
        }

        $Body = "<p class='MsoNormal' style='margin-bottom: 12.0pt; background: white;'><span style='font-size: 12pt; font-family: arial, helvetica, sans-serif;'>If you are part of this placeholder, you have an application to validate during this patch window. Please check SNOW for the specific application needing validation also enclosed is a list for your convenience. Managers are getting this placeholder for awareness.<br><br>The Server team will kick off patches for all servers at $PatchStart. An email will be sent to all applicable validators plus IT Mgmt. for awareness.<br><br><strong>-</strong>&nbsp;All Validators will use the Patching tab in the Service Now Business Service to mark validation as complete. You will no longer email the Service Desk. See attached &ldquo;IT Reference Sheet Prod Patching Validation in Business Service&rdquo; document for reference at the end of this email.</span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white;'><span style='font-size: 12pt; font-family: arial, helvetica, sans-serif;'><span style='color: black;'><span style='font-size: 12.0pt; font-family: Symbol; mso-fareast-font-family: Calibri; mso-fareast-theme-font: minor-latin; mso-bidi-font-family: Calibri; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA;'>&middot; </span></span>Validators for Automatic reboots should begin validation within <strong>two</strong>&nbsp;hours&nbsp;<strong>**or**</strong>&nbsp;after the competed email has been sent</span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt; font-family: arial, helvetica, sans-serif;'><span style='color: black;'><span style='font-size: 12.0pt; font-family: Symbol; mso-fareast-font-family: Calibri; mso-fareast-theme-font: minor-latin; mso-bidi-font-family: Calibri; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA;'>&middot; </span></span>Validators for Applications with Manual reboots will need to Lync server team member <strong>**or**</strong>&nbsp;join the conference bridge to verify if any additional patches are needed before validating.</span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt; font-family: arial, helvetica, sans-serif;'><span style='color: black;'><span style='font-size: 12.0pt; font-family: Symbol; mso-fareast-font-family: Calibri; mso-fareast-theme-font: minor-latin; mso-bidi-font-family: Calibri; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA;'>&middot; </span></span>Validators that encounter any issues should dial-in to the conference bridge or Service Desk.</span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt; font-family: arial, helvetica, sans-serif;'><span style='color: black;'><span style='font-size: 12.0pt; font-family: Symbol; mso-fareast-font-family: Calibri; mso-fareast-theme-font: minor-latin; mso-bidi-font-family: Calibri; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA;'>&middot; </span></span>If anyone is missed that should receive this placeholder please forward this invite to them.</span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt;'><strong><span style='font-family: 'Calibri',sans-serif;'>-----</span></strong></span></p>
        <p class='MsoNormal' style='margin-bottom: 12.0pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt;'><strong><em><u>Join Zoom Meeting:</u></em></strong></span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt;'><a href='https://zoom.us/my/somewhere'><span style='color: #337ab7;'>https://zoom.us/my/somewhere</span></a></span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt;'><strong><span style='font-family: 'Calibri',sans-serif;'>---------------------------------------------------</span></strong></span></p>
        <p class='MsoNormal' style='margin-bottom: 7.5pt; background: white; box-sizing: border-box; font-variant-ligatures: normal; font-variant-caps: normal; orphans: 2; text-align: start; widows: 2; -webkit-text-stroke-width: 0px; text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; word-spacing: 0px;'><span style='font-size: 12pt;'><em><span style='font-family: 'Calibri',sans-serif;'>IT Core Infrastructure Operations<br></span></em></span><span style='font-size: 12pt;'><strong><br><span style='font-family: 'Calibri',sans-serif;'>Information Technology</span></strong></span></p>
        <p class='MsoNormal'>&nbsp;</p>
        <p class='MsoNormal'><span style='font-size: 12pt;'>&nbsp;</span></p>"      
        
        $Mail.Body = $Body
        $Mail.Subject = $vCalTitle

        #String Builder Fun
        $sb = [System.Text.StringBuilder]::new()
        $sb.AppendLine("BEGIN:VCALENDAR") | Out-Null
        $sb.AppendLine("PRODID:-//Schedule a Meeting") | Out-Null
        $sb.AppendLine("VERSION:2.0") | Out-Null
        $sb.AppendLine("METHOD:REQUEST") | Out-Null

        #Configuration for correct Timezone
        $sb.AppendLine("BEGIN:VTIMEZONE") | Out-Null
        $sb.AppendLine("TZID:Central Standard Time") | Out-Null
        $sb.AppendLine("BEGIN:STANDARD") | Out-Null
        $sb.AppendLine("DTSTART:16011104T020000") | Out-Null
        $sb.AppendLine("RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11") | Out-Null
        $sb.AppendLine("TZOFFSETFROM:-0500") | Out-Null
        $sb.AppendLine("TZOFFSETTO:-0600") | Out-Null
        $sb.AppendLine("END:STANDARD") | Out-Null
        $sb.AppendLine("BEGIN:DAYLIGHT") | Out-Null
        $sb.AppendLine("DTSTART:16010311T020000") | Out-Null
        $sb.AppendLine("RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3") | Out-Null
        $sb.AppendLine("TZOFFSETFROM:-0600") | Out-Null
        $sb.AppendLine("TZOFFSETTO:-0500") | Out-Null
        $sb.AppendLine("END:DAYLIGHT") | Out-Null
        $sb.AppendLine("END:VTIMEZONE") | Out-Null

        #Start vEvent
        $sb.AppendLine("BEGIN:VEVENT") | Out-Null

        #Attendees
        $sb.AppendLine([String]::Format("{0}",$($Attendees).Trim())) | Out-Null

        $sb.AppendLine([String]::Format("DTSTART:{0:yyyyMMddTHHmmssZ}", $StartTime)) | Out-Null
        $sb.AppendLine([String]::Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", [datetime]::Now)) | Out-Null
        $sb.AppendLine([String]::Format("DTEND:{0:yyyyMMddTHHmmssZ}", $EndTime)) | Out-Null
        $sb.AppendLine("LOCATION: " + "+1123456789\,\,123456789# US (Chicago) /  https://zoom.us/my/somewhere") | Out-Null
        $sb.AppendLine([String]::Format("UID:{0}", $(New-Guid).Guid)) | Out-Null
        $sb.AppendLine([String]::Format("DESCRIPTION:{0}", $Mail.Body)) | Out-Null
        $sb.AppendLine([String]::Format("X-ALT-DESC;FMTTYPE=text/html:{0}", $Mail.Body)) | Out-Null
        $sb.AppendLine([String]::Format("SUMMARY:{0}", $vCalTitle)) | Out-Null
        $sb.AppendLine([String]::Format("SUMMARY:{0}", $vCalTitle)) | Out-Null
        $sb.AppendLine([String]::Format("ORGANIZER:MAILTO:{0}", $Mail.From.Address)) | Out-Null

        $sb.AppendLine("BEGIN:VALARM") | Out-Null
        $sb.AppendLine("TRIGGER:-PT15M") | Out-Null
        $sb.AppendLine("ACTION:DISPLAY") | Out-Null
        $sb.AppendLine("DESCRIPTION:Reminder") | Out-Null
        $sb.AppendLine("END:VALARM") | Out-Null
        $sb.AppendLine("END:VEVENT") | Out-Null
        $sb.AppendLine("END:VCALENDAR") | Out-Null

        #Configure mimetype to support our Calendar within the body of the email
        $contype = New-Object System.Net.Mime.ContentType("text/calendar")
        $contype.Parameters.Add("method", "REQUEST")
        $contype.Parameters.Add("name", "Test.ics");
        $avCal = [Net.Mail.AlternateView]::CreateAlternateViewFromString($sb.ToString(),$contype)

        #Configure mimetype to support HTML in the body of the email
        $EnableHTML = New-Object System.Net.Mime.ContentType("text/html")
        $aHTML = [Net.Mail.AlternateView]::CreateAlternateViewFromString($Body.ToString(),$EnableHTML)

        $Mail.AlternateViews.Add($aHTML)
        $Mail.AlternateViews.Add($avCal)

        $File1 = New-Object System.Net.Mail.LinkedResource($ServerList, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        $File1.ContentType.Name = "Filtered-Spreadsheet.xlsx"
        $aHTML.LinkedResources.Add($File1)

        $File2 = New-Object System.Net.Mail.LinkedResource($ValidationFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        $File2.ContentType.Name = "Patching Validation in Business Service.docx"
        $aHTML.LinkedResources.Add($File2)
        
        #=======================================

        $SendTo = (($Attendees -split ("\s+")) -split ("mailto:")) | select-object | Where-Object {$_ -like "*@*"}

        foreach($To in $SendTo)
        {
            #$To
            $Mail.To.Add($To)
        }  

        #DEBUGGING
        #$Mail.To
        #Write-Output "`n-----------------------------------------`n"

        #Send Email
		#(Comment the line below if you are testing so the email doesn't get sent out)
        $SmtpClient.Send($Mail)
        
        #Clear the "To Addresses" each iteration
        $Mail.To.Clear()

        #Free Resources
        $aHTML.LinkedResources.Remove($File1) | Out-Null
        $aHTML.LinkedResources.Remove($File2) | Out-Null

        $Mail.AlternateViews.Remove($aHTML) | Out-Null
        $Mail.AlternateViews.Remove($avCal) | Out-Null
        
        #DEBUG
        #$To
        #$Mail
        #$vCalTitle
    }
}

#Clear resources
$Mail.Dispose()
$SmtpClient.Dispose()
