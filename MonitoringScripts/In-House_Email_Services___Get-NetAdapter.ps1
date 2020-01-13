Remove-Item "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Get-NetAdapter_Report.htm"

cls

$Script_Start_Time = Get-Date
write-host $Script_Start_Time

$ServerList = Get-ExchangeServer | where {$_.Name -like "EIS-*-EP0*"} | sort Name | select Name
#$ServerList = Get-ExchangeServer | where {$_.Name -like "EIS-LS2-EP06143" -or $_.Name -like "EIS-LS2-EP06144" -or $_.Name -like "EIS-LS2-EP06145"} | sort Name | select Name

$Result = @()
$Results = @()
$ReportDate = Get-Date
$WarningEvents = 0
$ErrorEvents = 0

$Outputreport = "<HTML><TITLE> In-House Email Services - Get-NetAdapter Report </TITLE>
                 <BODY background-color:peachpuff>
                 <font color =""#99000"" face=""Microsoft Tai le"">
                 <H2> In-House Email Services - Get-NetAdapter Report - $ReportDate</H2></font>
                 <Table border=2 cellpadding=5 cellspacing=0>
                 <TR bgcolor=Aqua align=center>
                   <TD><B>Server Name</B></TD>
                   <TD><B>Adapter Name</B></TD>
                   <TD><B>InterfaceDescription</B></TD>
                   <TD><B>ifIndex</B></TD>
                   <TD><B>Status</B></TD>
                   <TD><B>MacAddress</B></TD>
                   <TD><B>LinkSpeed</B></TD></TR>"

foreach ($Server in $ServerList) {
    $NetAdapterInfo = $null
    $NetAdapterInfo = Get-NetAdapter -CimSession $Server.Name | where {$_.Status -notlike "Not Present"} | Select SystemName,Name,InterfaceDescription,ifIndex,Status,MacAddress,LinkSpeed
    
    foreach ($NetAdapter in $NetAdapterInfo)

        {$Result = $null
         $Result += [PSCustomObject] @{ 
        
        ServerName = $NetAdapter.SystemName
        AdapterName = $NetAdapter.Name
        InterfaceDescription = $NetAdapter.InterfaceDescription
        ifIndex = $NetAdapter.ifIndex
        Status = $NetAdapter.Status
        MacAddress = $NetAdapter.MacAddress
        LinkSpeed = $NetAdapter.LinkSpeed}

        $Results += $Result}}
                                      
        Foreach($Entry in $Results) 
    
            {  if ($Entry.Status -notlike "Up"){
                                      
                $Outputreport += "<TR bgcolor=Red>"
                $ErrorEvents ++}

                elseif($Entry.AdapterName -like "DAG Replication"){
                
                $Outputreport += "<TR bgcolor=LightBlue>"}
                    
                else
            {$Outputreport += "<TR>"                      }
             $Outputreport += "<TD>$($Entry.Servername)</TD><TD align=center>$($Entry.AdapterName)</TD><TD align=center>$($Entry.InterfaceDescription)</TD><TD align=center>$($Entry.ifIndex)</TD><TD align=center>$($Entry.Status)</TD><TD align=center>$($Entry.MacAddress)</TD><TD align=center>$($Entry.LinkSpeed)</TD></TR>" 
            }
        
   $Outputreport += "</Table></BODY></HTML>"

$Outputreport | out-file C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Get-NetAdapter_Report.htm
    
$Script_Stopped_Time = Get-Date
write-host "In-House Email Services - Get-NetAdapter Report completed: " $Script_Stopped_Time

if ($ErrorEvents -gt 0)
    {$EmailPriority = "High"}

elseif ($WarningEvents -gt 0)
    {$EmailPriority = "Normal"}

else
    {$EmailPriority = "Low"}

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Get-NetAdapter_Report.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "Kwok-Fai.HA@tdc.forces.gc.ca"

$Subject = "In-House Email Services - Get-NetAdapter Report - $ReportDate"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}