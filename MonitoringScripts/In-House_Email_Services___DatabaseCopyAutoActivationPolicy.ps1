Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___DatabaseCopyAutoActivationPolicy.htm"

cls

$Result = @()
$ReportDate = Get-Date
$WarningEvents = 0
$ErrorEvents = 0

$Outputreport = "<HTML><TITLE> In-House Email Services - Exchange DatabaseCopyAutoActivationPolicy Report </TITLE>
                 <BODY background-color:peachpuff>
                 <font color =""#99000"" face=""Microsoft Tai le"">
                 <H2> In-House Email Services - Exchange DatabaseCopyAutoActivationPolicy Report - $ReportDate</H2></font>
                 <Table border=2 cellpadding=5 cellspacing=0>
                 <TR bgcolor=CadetBlue align=center>
                   <TD><B>Name</B></TD>
                   <TD><B>DatabaseAvailabilityGroup</B></TD>
                   <TD><B>MAPIEncryptionRequired</B></TD>
                   <TD><B>AutoDatabaseMountDial</B></TD>
                   <TD><B>DatabaseCopyAutoActivationPolicy</B></TD></TR>"

$Script_Start_Time = Get-Date
Write-Host $Script_Start_Time -ForegroundColor Green
Write-Host ""

Write-Host "Checking Exchange DatabaseCopyAutoActivationPolicy.  Please wait (approx. 3 minutes)..." -ForegroundColor Magenta

$DatabaseCopyAutoActivationPolicy_Information = Get-MailboxServer | where {($_.name -like "EIS-*-EP0*") -or ($_.Name -like "IMG-*") -or ($_.Name -like "DCS-*")} | sort Name

foreach ($DatabaseCopyAutoActivationPolicyInfo in $DatabaseCopyAutoActivationPolicy_Information){

          if($DatabaseCopyAutoActivationPolicyInfo.DatabaseCopyAutoActivationPolicy -notlike "Unrestricted") 

            {$Outputreport += "<TR bgcolor=Aqua>"
             $ErrorEvents ++}
                 
          else
            {$Outputreport += "<TR>"                      }
             $Outputreport += "
             <TD>$($DatabaseCopyAutoActivationPolicyInfo.Name)</TD>
             <TD align=center>$($DatabaseCopyAutoActivationPolicyInfo.DatabaseAvailabilityGroup)</TD><TD align=center>$($DatabaseCopyAutoActivationPolicyInfo.MAPIEncryptionRequired)</TD><TD align=center>$($DatabaseCopyAutoActivationPolicyInfo.AutoDatabaseMountDial)</TD><TD align=center>$($DatabaseCopyAutoActivationPolicyInfo.DatabaseCopyAutoActivationPolicy)</TD></TR>"
            }

$Outputreport += "</Table></BODY></HTML>"
$Outputreport | out-file "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___DatabaseCopyAutoActivationPolicy.htm"

$Script_Stopped_Time = Get-Date
Write-Host ""
write-host "Exchange Exchange DatabaseCopyAutoActivationPolicy check completed: " $Script_Stopped_Time -ForegroundColor Green

if ($ErrorEvents -lt 14)
    {$EmailPriority = "High"}

else
    {$EmailPriority = "Low"}

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___DatabaseCopyAutoActivationPolicy.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - Exchange DatabaseCopyAutoActivationPolicy Report"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}