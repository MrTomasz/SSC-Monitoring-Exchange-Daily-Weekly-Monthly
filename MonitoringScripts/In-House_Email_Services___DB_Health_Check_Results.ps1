Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___DB_Health.htm"

cls

$Result = @()
$ReportDate = Get-Date
$WarningEvents = 0
$ErrorEvents = 0

$Outputreport = "<HTML><TITLE> In-House Email Services Database Health Check Results </TITLE>
                 <BODY background-color:peachpuff>
                 <font color =""#99000"" face=""Microsoft Tai le"">
                 <H2> In-House Email Services - Database Health Check Results - $ReportDate</H2></font>
                 <Table border=2 cellpadding=5 cellspacing=0>
                 <TR bgcolor=CadetBlue align=center>
                   <TD><B>Name</B></TD>
                   <TD><B>Status</B></TD>
                   <TD><B>CopyQueueLength (>50)</B></TD>
                   <TD><B>ReplayQueueLength (>700)</B></TD>
                   <TD><B>LastInspectedLogTime</B></TD>
                   <TD><B>ContentIndexState</B></TD></TR>"

$Script_Start_Time = Get-Date
Write-Host $Script_Start_Time -ForegroundColor Green
Write-Host ""

Write-Host "Checking Exchange Databases health.  Please wait (approx. 1 minute)..." -ForegroundColor Magenta

$DB_Info = Get-MailboxServer | where {($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")} | Get-MailboxDatabaseCopyStatus | sort status

foreach ($DB in $DB_Info){

          if(($DB.Status -like "Failed*") -or ($DB.Status -like "ServiceDown") -or ($DB.ContentIndexState -like "Failed"))

            {$Outputreport += "<TR bgcolor=Red>"
             $ErrorEvents ++}
              
          elseif(($DB.CopyQueueLength -gt 50) -or ($DB.ReplayQueueLength -gt 700))
                    
            {$Outputreport += "<TR bgcolor=Yellow>" 
             $WarningEvents ++}

          else
            {$Outputreport += "<TR>"                      }
             $Outputreport += "<TD>$($DB.Name)</TD><TD align=center>$($DB.Status)</TD><TD align=center>$($DB.CopyQueueLength)</TD><TD align=center>$($DB.ReplayQueueLength)</TD><TD align=center>$($DB.LastInspectedLogTime)</TD><TD align=center>$($DB.ContentIndexState)</TD></TR>"
            }

$Outputreport += "</Table></BODY></HTML>"
$Outputreport | out-file "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___DB_Health.htm"

$Script_Stopped_Time = Get-Date
Write-Host ""
write-host "Exchange Databases health check completed: " $Script_Stopped_Time -ForegroundColor Green

if ($ErrorEvents -gt 0)
    {$EmailPriority = "High"}

elseif ($WarningEvents -gt 0)
    {$EmailPriority = "Normal"}
else
    {$EmailPriority = "Low"}

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___DB_Health.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - Database Health Check Results"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}