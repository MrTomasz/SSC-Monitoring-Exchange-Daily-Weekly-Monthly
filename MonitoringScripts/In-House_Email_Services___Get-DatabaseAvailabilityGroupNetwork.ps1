Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Get-DatabaseAvailabilityGroupNetwork_Report.htm"

cls

$Script_Start_Time = Get-Date
write-host $Script_Start_Time

#$DAG_Networks = Get-DatabaseAvailabilityGroupNetwork | where {$_.Identity -like "DEMS-DAG-OTG*"} | sort Identity.

$DAGs = Get-DatabaseAvailabilityGroup | where {$_.Name -like "NAT*" -or $_.Name -like "SFK*"}
$DAGs_Networks = foreach ($DAG in $DAGs) {Get-DatabaseAvailabilityGroupNetwork -Identity $DAG.Name}

#$DAG_Networks = Get-DatabaseAvailabilityGroupNetwork | where {$_.Identity -like "DEMS*"} | sort Identity
#$ServerList = Get-ExchangeServer | where {$_.Name -like "EIS-LS2-EP06143" -or $_.Name -like "EIS-LS2-EP06144" -or $_.Name -like "EIS-LS2-EP06145"} | sort Name | select Name

$Result = @()
$Results = @()
$ReportDate = Get-Date
$WarningEvents = 0
$ErrorEvents = 0

$Outputreport = "<HTML><TITLE> In-House Email Services - Get-DatabaseAvailabilityGroupNetwork Report</TITLE>
                 <BODY background-color:peachpuff>
                 <font color =""#99000"" face=""Microsoft Tai le"">
                 <H2> In-House Email Services - Get-DatabaseAvailabilityGroupNetwork Report - $ReportDate</H2></font>
                 <Table border=2 cellpadding=5 cellspacing=0>
                 <TR bgcolor=Aqua align=center>
                   <TD><B>Identity</B></TD>
                   <TD><B>MapiAccessEnabled</B></TD>
                   <TD><B>ReplicationEnabled</B></TD>
                   <TD><B>Subnets</B></TD></TR>"
                   
foreach ($Connection in $DAGs_Networks) {
    
         #$Result = $null
         $Result += [PSCustomObject] @{ 
        
        Identity = $Connection.Identity
        MapiAccessEnabled = $Connection.MapiAccessEnabled
        ReplicationEnabled = $Connection.ReplicationEnabled
        Subnets = $Connection.Subnets}
                    
         $Results += $Result
         $Result = $null}
                                                
Foreach ($Entry in $Results) 
    
            {   
                if (($Entry.Identity -like "NAT*" -or $Entry.Identity -like "SFK*") -and $Entry.Identity -like "*MAPI*" -and $Entry.ReplicationEnabled -eq "True"){
                $Outputreport += "<TR bgcolor=Red>"
                $ErrorEvents ++}

                elseif ($Entry.Identity -like "*Replication*"){
                $Outputreport += "<TR bgcolor=LightBlue>"}
                    
                else {
                $Outputreport += "<TR>" }
                                  
             $Outputreport += "<TD>$($Entry.Identity)</TD><TD align=center>$($Entry.MapiAccessEnabled)</TD><TD align=center>$($Entry.ReplicationEnabled)</TD><TD align=center>$($Entry.Subnets)</TD></TR>" 
            }
        
   $Outputreport += "</Table></BODY></HTML>"

$Outputreport | out-file C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Get-DatabaseAvailabilityGroupNetwork_Report.htm
    
$Script_Stopped_Time = Get-Date
write-host "In-House Email Services - Get-DatabaseAvailabilityGroupNetwork Report completed: " $Script_Stopped_Time

if ($ErrorEvents -gt 0)
    {$EmailPriority = "High"}

elseif ($WarningEvents -gt 0)
    {$EmailPriority = "Normal"}

else
    {$EmailPriority = "Low"}

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Get-DatabaseAvailabilityGroupNetwork_Report.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "Kwok-Fai.HA@tdc.forces.gc.ca"

$Subject = "In-House Email Services - Get-DatabaseAvailabilityGroupNetwork Report - $ReportDate"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}