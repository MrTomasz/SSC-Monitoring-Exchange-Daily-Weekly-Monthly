Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___ActiveSync_Health.htm"

cls

$Result = @()
$ReportDate = Get-Date
$WarningEvents = 0
$ErrorEvents = 0

$Outputreport = "<HTML><TITLE> In-House Email Services - ActiveSync Health Check Results </TITLE>
                 <BODY background-color:peachpuff>
                 <font color =""#99000"" face=""Microsoft Tai le"">
                 <H2> In-House Email Services - ActiveSync Health Check Results $ReportDate</H2></font>
                 <Table border=2 cellpadding=5 cellspacing=0>
                 <TR bgcolor=CadetBlue align=center>
                   <TD><B>Server</B></TD>
                   <TD><B>State</B></TD>
                   <TD><B>Name</B></TD>
                   <TD><B>TargetResource</B></TD>
                   <TD><B>HealthSetNamee</B></TD>
                   <TD><B>AlertValue</B></TD>
                   <TD><B>ServerComponent</B></TD></TR>"

$Script_Start_Time = Get-Date
Write-Host $Script_Start_Time -ForegroundColor Green
Write-Host ""

Write-Host "Checking Exchange ActiveSync health.  Please wait (approx. 3 minutes)..." -ForegroundColor Magenta

$ActiveSync_Information = Get-ExchangeServer | where {$_.name -like "EIS-*-EP0*"} | Get-ServerHealth | where {$_.HealthSetName -like "ActiveSync*"} | sort Server

foreach ($ActiveSyncInfo in $ActiveSync_Information){

          if($ActiveSyncInfo.AlertValue -notlike "Healthy") 

            {$Outputreport += "<TR bgcolor=Red>"
             $ErrorEvents ++}
                 
          else
            {$Outputreport += "<TR>"                      }
             $Outputreport += "<TD>$($ActiveSyncInfo.Server)</TD><TD align=center>$($ActiveSyncInfo.CurrentHealthSetState)</TD><TD align=center>$($ActiveSyncInfo.Name)</TD><TD align=center>$($ActiveSyncInfo.TargetResource)</TD><TD align=center>$($ActiveSyncInfo.HealthSetName)</TD><TD align=center>$($ActiveSyncInfo.AlertValue)</TD><TD align=center>$($ActiveSyncInfo.ServerComponentName)</TD></TR>"
            }

$Outputreport += "</Table></BODY></HTML>"
$Outputreport | out-file "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___ActiveSync_Health.htm"

$Script_Stopped_Time = Get-Date
Write-Host ""
write-host "Exchange ActiveSync health check completed: " $Script_Stopped_Time -ForegroundColor Green

if ($ErrorEvents -gt 0)
    {$EmailPriority = "High"}

elseif ($WarningEvents -gt 0)
    {$EmailPriority = "Normal"}

else
    {$EmailPriority = "Low"}

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___ActiveSync_Health.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - ActiveSync Health Check Results"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}