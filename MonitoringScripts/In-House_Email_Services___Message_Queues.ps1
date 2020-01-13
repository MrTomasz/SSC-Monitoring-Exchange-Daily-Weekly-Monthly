Remove-Item In-House_Email_Services___Message_Queues.htm

CLS

$Results = @()
$ReportDate = Get-Date

$Script_Start_Time = Get-Date
write-host "In-House Email Services message queues check started: " $Script_Start_Time -ForegroundColor Magenta

$Results = Get-ExchangeServer | where {(($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")) -and (($_.ServerRole -like "*HubTransport*") -or ($_.ServerRole -like "Mailbox*"))}`
 | sort Name | get-queue | ? {-not ($_.DeliveryType -like "Shadow*") -and $_.MessageCount -ge 0} `
 | sort MessageCount -Descending | Select Identity,DeliveryType,Status,MessageCount,NextHopDomain,LastRetryTime,NextRetryTime,LastError

$Script_Stopped_Time = Get-Date
write-host "In-House Email Services message queues check completed: " $Script_Stopped_Time -ForegroundColor Green

$Outputreport = "<HTML><TITLE> In-House Email Services - Message Queue Results </TITLE>
                     <BODY background-color:peachpuff>
                     <font color =""#99000"" face=""Microsoft Tai le"">
                     <H2> In-House Email Services - Message Queue Results $ReportDate</H2></font>
                     <Table border=2 cellpadding=5 cellspacing=0>
                     <TR bgcolor=MediumPurple align=center>
                       <TD><B>IDENTITY</B></TD>
                       <TD><B>DELIVERY TYPE</B></TD>
                       <TD><B>STATUS</B></TD>
                       <TD><B>MESSAGE COUNT</B></TD>
                       <TD><B>NEXT HOP DOMAIN</B></TD>
                       <TD><B>LAST RETRY TIME</B></TD>
                       <TD><B>NEXT RETRY TIME</B></TD>
                       <TD><B>LAST ERROR</B></TD></TR>"
                 
    Foreach($Entry in $Results) 
    
        {        
           $Outputreport += "<TR>"                      
           $Outputreport += "<TD>$($Entry.Identity)</TD>`
                            <TD align=center>$($Entry.DeliveryType)</TD>`
                            <TD align=center>$($Entry.Status)</TD>`
                            <TD align=center>$($Entry.MessageCount)</TD>`
                            <TD align=center>$($Entry.NextHopDomain)</TD>`
                            <TD align=center>$($Entry.LastRetryTime)</TD>`
                            <TD align=center>$($Entry.NextRetryTime)</TD>`
                            <TD align=center>$($Entry.LastError)</TD></TR>" 
        }
     $Outputreport += "</Table></BODY></HTML>"

$Outputreport | out-file C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Message_Queues.htm
#Invoke-Expression C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\DEMS_MessageQueues.htm

##Send email functionality from below line, use it if you want

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Message_Queues.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - Message Queue Results - $ReportDate"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}