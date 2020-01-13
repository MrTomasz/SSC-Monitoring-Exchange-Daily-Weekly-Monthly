Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___ServerComponentState.htm"

cls

$Result = @()
$ReportDate = Get-Date
$WarningEvents = 0
$ErrorEvents = 0

$Outputreport = "<HTML><TITLE> In-House Email Services - Exchange 2016 ServerComponentState Results </TITLE>
                 <BODY background-color:peachpuff>
                 <font color =""#99000"" face=""Microsoft Tai le"">
                 <H2> In-House Email Services - Exchange 2016 ServerComponentState Results - $ReportDate</H2></font>
                 <Table border=2 cellpadding=5 cellspacing=0>
                 <TR bgcolor=CadetBlue align=center>
                   <TD><B>Server</B></TD>
                   <TD><B>ServerComponent</B></TD>
                   <TD><B>State</B></TD></TR>"

$Script_Start_Time = Get-Date
Write-Host $Script_Start_Time -ForegroundColor Green
Write-Host ""

Write-Host "Checking Exchange 2016 ServerComponentState.  Please wait (approx. 3 minutes)..." -ForegroundColor Magenta

$ServerComponentState_Information = Get-ExchangeServer | where {$_.name -like "EIS-*-EP0*"} | Get-ServerComponentState | sort Identity

foreach ($ServerComponentStateInfo in $ServerComponentState_Information){

          if(($ServerComponentStateInfo.State -notlike "Active" -and $ServerComponentStateInfo.Component -notlike "ForwardSyncDaemon") -and ($ServerComponentStateInfo.State -notlike "Active" -and $ServerComponentStateInfo.Component -notlike "ProvisioningRps")) 

            {$Outputreport += "<TR bgcolor=Red>"
             $ErrorEvents ++}

          elseif(($ServerComponentStateInfo.State -notlike "Active" -and $ServerComponentStateInfo.Component -like "ForwardSyncDaemon") -or ($ServerComponentStateInfo.State -notlike "Active" -and $ServerComponentStateInfo.Component -like "ProvisioningRps")) 
              
            {$Outputreport += "<TR bgcolor=Yellow>"}
                 
          else
            {$Outputreport += "<TR>"                      }
             $Outputreport += "<TD>$($ServerComponentStateInfo.Identity)</TD><TD align=center>$($ServerComponentStateInfo.Component)</TD><TD align=center>$($ServerComponentStateInfo.State)</TD></TR>"
            }

$Outputreport += "</Table></BODY></HTML>"
$Outputreport | out-file "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___ServerComponentState_Health.htm"

$Script_Stopped_Time = Get-Date
Write-Host ""
write-host "Exchange 2016 ServerComponentState check completed: " $Script_Stopped_Time -ForegroundColor Green

if ($ErrorEvents -gt 0)
    {$EmailPriority = "High"}

elseif ($WarningEvents -gt 0)
    {$EmailPriority = "Normal"}

else
    {$EmailPriority = "Low"}

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___ServerComponentState_Health.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - Exchange 2016 ServerComponentState Health Check Results"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}