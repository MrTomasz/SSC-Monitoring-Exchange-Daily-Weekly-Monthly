Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Test-ReplicationHealth.txt"

cls

$Script_Start_Time = Get-Date
write-host $Script_Start_Time

#$MailboxServers = Get-MailboxServer | where {($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")} | sort name
$MailboxServers = Get-MailboxServer | where {$_.name -like "EIS-*-EP0*" } | sort name

foreach ($MailboxServer in $MailboxServers) {
$MailboxServer | Test-ReplicationHealth | ft -AutoSize | Out-File "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Test-ReplicationHealth.txt" -Append}
#$MailboxServer | Test-ReplicationHealth | where {$_.Result -notlike "Passed"} | fl Server,Check,Result,Error | Out-File "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Test-ReplicationHealth.txt" -Append}

$Script_Stopped_Time = Get-Date
write-host "In-House Email Services Exchange Replication Health check completed: " $Script_Stopped_Time

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Test-ReplicationHealth.txt"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "Kwok-Fai.HA@tdc.forces.gc.ca"

$Subject = "In-House Email Services - Exchange Replication Health Check Results"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}