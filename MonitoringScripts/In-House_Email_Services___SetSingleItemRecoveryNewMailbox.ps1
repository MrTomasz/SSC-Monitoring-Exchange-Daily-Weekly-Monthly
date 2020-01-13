# In-House_Email_Services___SetSingleItemRecoveryNewMailbox.ps1
# Updated on 2019 10 01

cls

Write-Host "Start Time: " (Get-Date) -ForegroundColor Green
Write-Host ""

$DBs_Count = 0
$DBs_Progress = 1
$ALL_NewMailboxes = @()
$DB_NewMailboxes = @()

$DB_Mailboxes_Count = 0
$DB_NewMailboxes_Count = 0
$ALL_Mailboxes_Count = 0
$ALL_NewMailboxes_Count = 0

$DBs = Get-MailboxDatabase | where {$_.Server -like "EIS-DCB-EP061*" -or $_.Server -like "EIS-LS2-EP061*" -or $_.Server -like "EIS-SFK-EP062*"} | sort Name
#$DBs = Get-MailboxDatabase -IncludePreExchange2013 | where {$_.Server -like "EIS-DCB-E*" -or $_.Server -like "EIS-LS2-E*" -or $_.Server -like "IMG-DCB-E00*" -or $_.Server -like "IMG-ESQ-E00*" -or $_.Server -like "IMG-OTG-E00*" } | sort Name

$DBs_Count = ($DBs).Count

Write-Host "  Found: " $DBs_Count "databases" -ForegroundColor Magenta
Write-Host ""
Write-Host "-------------------------------------------------------------------------------------"
Write-Host ""

foreach ($DB in $DBs)
{ 

Write-Host "Start Time: " (Get-Date) -ForegroundColor Green
Write-Host ""
Write-Host "  Database: " $DBs_Progress "of" $DBs_Count -ForegroundColor Magenta
Write-Host ""
Write-Host "Working on database: " $DB -ForegroundColor Cyan
Write-Host ""

$DB_Mailboxes = Get-Mailbox -Database $DB -ResultSize Unlimited
$DB_Mailboxes_Count = ($DB_Mailboxes).Count

$ALL_Mailboxes_Count += $DB_Mailboxes_Count

#Normal query to get new mailboxes from last day
$DB_NewMailboxes = $DB_Mailboxes | Where { $_.WhenCreated -gt (Get-Date).AddDays(-1) }

#Initial query to get all mailboxes on 2016
#$DB_NewMailboxes = $DB_Mailboxes 

$DB_NewMailboxes_Count = ($DB_NewMailboxes).Count

Write-Host "Found:  " $DB_Mailboxes_Count "mailboxes in db"
Write-Host "Found:  " $DB_NewMailboxes_Count "new mailboxes"
Write-Host ""
Write-Host "Setting SingleItemRecoveryEnabled and UseDatabaseQuotaDefaults to True for" $DB_NewMailboxes_Count "new mailboxes" -ForegroundColor Cyan
Write-Host ""

foreach ($DB_NewMailbox in $DB_NewMailboxes)
{
    Set-Mailbox $DB_NewMailbox -SingleItemRecoveryEnabled $True -UseDatabaseQuotaDefaults $True

    $DB_NewMailbox

$ALL_NewMailboxes += $DB_NewMailbox

}

Write-Host ""
Write-Host "End Time: " (Get-Date) -ForegroundColor Red
Write-Host ""
Write-Host "-------------------------------------------------------------------------------------"
Write-Host ""

$DBs_Progress = $DBs_Progress + 1

}

$ALL_NewMailboxes_Count = ($ALL_NewMailboxes).Count

Write-Host "Found:  " $ALL_Mailboxes_Count "mailboxes in total"
Write-Host "Found:  " $ALL_NewMailboxes_Count "new mailboxes in total"

$Report = $ALL_NewMailboxes | Select-Object ServerName,Database,DisplayName,IsMailboxEnabled,SingleItemRecoveryEnabled,RetainDeletedItemsFor,UseDatabaseRetentionDefaults,RetainDeletedItemsUntilBackup,LitigationHoldEnabled,RetentionHoldEnabled,UseDatabaseQuotaDefaults,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,OrganizationalUnit,UserPrincipalName,Alias,SamAccountName,PrimarySmtpAddress,DistinguishedName,Identity,WhenCreated
$Report | ft -AutoSize

###############################################################################

#Create variable for log date

$LogDate = Get-Date -f yyyy-MM-dd_HH-mm-ss

#Export report to CSV file 

$Report | Export-Csv -path "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\Reports\In-House_Email_Services___SingleItemRecoveryNewMailbox_$logDate.csv" -NoTypeInformation

#Send report using the exchange email module

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\Reports\In-House_Email_Services___SingleItemRecoveryNewMailbox_$logDate.csv"

#send-mailmessage -from Admin@forces.gc.ca -to DEMSCEMOPS@forces.gc.ca -subject "TDC DEMS Network Info" -SMTPServer IMG-LSL-EP00895 -Attachments $EMailAttachment

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - Set SingleItemRecovery on $ALL_NewMailboxes_Count New Mailboxes"
$Body = "Processed $ALL_Mailboxes_Count mailboxes in total."

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -body $Body -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -body $Body -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -body $Body -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}

###############################################################################