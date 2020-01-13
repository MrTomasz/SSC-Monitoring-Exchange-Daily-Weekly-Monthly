Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Test-ServiceHealth.txt"

cls

$Script_Start_Time = Get-Date
write-host $Script_Start_Time

$ExchangeServers = Get-ExchangeServer | where {($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")} | sort name

foreach ($ExchangeServer in $ExchangeServers) {
$ExchangeServer | Test-ServiceHealth | ft $ExchangeServer.Name,Role,RequiredServicesRunning,ServicesNotRunning -AutoSize | Out-File "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Test-ServiceHealth.txt" -Append}

$Script_Stopped_Time = Get-Date
write-host "In-House Email Services Exchange Service Health check completed: " $Script_Stopped_Time

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Test-ServiceHealth.txt"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "Kwok-Fai.HA@tdc.forces.gc.ca"

$Subject = "In-House Email Services - Exchange Services Health Check Results"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}