CLS

#$ServerList = Get-Content "c:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\NITL_DEMS_Exchange_Servers_incl_2016_minus_ships_and_x.txt"
$ServerList = Get-ExchangeServer | where {(($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")) -and (($_.ServerRole -like "*HubTransport*") -or ($_.ServerRole -like "Mailbox*"))} | sort name | Select Name


$Results = @()
$ReportDate = Get-Date
$ErrorEvents = 0

ForEach($ComputerName in $ServerList) 
{  

$Data1 = Get-WMIObject -ComputerName $ComputerName.Name -Class win32_service | Where {$_.Name -EQ "SepMasterService"} | Select-Object @{Label='Server';Expression={$ComputerName.Name}},Name,StartMode,State,Status
$Data2 = Get-WMIObject -ComputerName $ComputerName.Name -Class win32_service | Where {$_.Name -EQ "SMSMSE"} | Select-Object @{Label='Server';Expression={$ComputerName.Name}},Name,StartMode,State,Status

If ($Data2 -eq $null){
 
        $Result = [PSCustomObject] @{ 
        
        ServerName = $ComputerName.Name
        SEPServiceName = $Data1.Name
        StartMode = $Data1.StartMode
        State = $Data1.State
        Status = $Data1.Status }
        
        $Results += $Result
        } 
        
Else {         
        $Result = [PSCustomObject] @{ 
        
        ServerName = $ComputerName.Name        
        SEPServiceName = $Data1.Name
        StartMode = $Data1.StartMode
        State = $Data1.State
        Status = $Data1.Status} 

        $Results += $Result

        $Result = [PSCustomObject] @{ 
                   
        ServerName = $ComputerName.Name
        SEPServiceName = $Data2.Name
        StartMode = $Data2.StartMode
        State = $Data2.State
        Status = $Data2.Status}

        $Results += $Result
        }}
                     
$Outputreport = "<HTML><TITLE> In-House Email Services - Symantec Services Report </TITLE>
                     <BODY background-color:peachpuff>
                     <font color =""#99000"" face=""Microsoft Tai le"">
                     <H2> In-House Email Services - Symantec Services Report - $ReportDate</H2></font>
                     <Table border=2 cellpadding=5 cellspacing=0>
                     <TR bgcolor=Yellow align=center>
                       <TD><B> SERVER NAME </B></TD>
                       <TD><B> SYMANTEC SERVICE NAME </B></TD>
                       <TD><B> START MODE </B></TD>
                       <TD><B> STATE </B></TD>
                       <TD><B> STATUS </B></TD></TR>"                 
    
    Foreach($Entry in $Results) 
    
        {        
            if($Entry.State -ne "Running")
            {
                $Outputreport += "<TR bgcolor=red>"
                $ErrorEvents ++
            } 
            else
           {
           $Outputreport += "<TR>"                      }
           $Outputreport += "<TD>$($Entry.Servername)</TD><TD align=center>$($Entry.SEPServiceName)</TD><TD align=center>$($Entry.StartMode)</TD><TD align=center>$($Entry.State)</TD><TD align=center>$($Entry.Status)</TD></TR>" 
        }
     $Outputreport += "</Table></BODY></HTML>"

$Outputreport | out-file C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Symantec_Report.htm
#Invoke-Expression C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Symantec_Report.htm

##Send email functionality from below line, use it if you want

if ($ErrorEvents -gt 0)
    {$EmailPriority = "High"}
else
    {$EmailPriority = "Low"}
    
$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Symantec_Report.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"

$Subject = "In-House Email Services - Symantec Services Check Results - $ReportDate"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}