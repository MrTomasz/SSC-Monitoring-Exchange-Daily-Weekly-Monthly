CLS

#$ServerList = Get-Content "c:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\NITL_DEMS_Exchange_Servers_incl_2016_minus_ships_and_x.txt"
$ServerList = Get-MailboxServer | where {($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")} | sort Name

$Results = @()
$ReportDate = Get-Date

ForEach($Server in $ServerList)
{

$Result = [PSCustomObject] @{ 
        
        ServerName = $Server
        RPCUser = (Get-Counter "\MSExchange RpcClientAccess\User Count" -ComputerName $Server).CounterSamples[0].Cookedvalue
        RPCCon = (Get-Counter "\MSExchange RpcClientAccess\Connection Count" -ComputerName $Server).CounterSamples[0].Cookedvalue 
        OWA = (Get-Counter "\MSExchange OWA\Current Unique Users" -ComputerName $Server).CounterSamples[0].Cookedvalue
        EAS1 = (Get-Counter "\MSExchange ActiveSync\Current Requests" -ComputerName $Server).CounterSamples[0].Cookedvalue
        EAS2 = [math]::Truncate((Get-Counter "\MSExchange ActiveSync\Requests/sec" -ComputerName $Server).CounterSamples[0].Cookedvalue)
        W3SVC = (Get-Counter "\W3SVC_W3WP(_Total)\Active Requests" -ComputerName $Server).CounterSamples[0].Cookedvalue
        WEB1 = (Get-Counter "\Web Service(_Total)\Current Connections" -ComputerName $Server).CounterSamples[0].Cookedvalue
        WEB2 = (Get-Counter "\Web Service(_Total)\Maximum Connections" -ComputerName $Server).CounterSamples[0].Cookedvalue}
                
        $Results += $Result
}

$Outputreport = "<HTML><TITLE> In-House Email Services - Connections Report </TITLE>
                     <BODY background-color:peachpuff>
                     <font color =""#99000"" face=""Microsoft Tai le"">
                     <H2> In-House Email Services - Connections Report $ReportDate</H2></font>
                     <Table border=2 cellpadding=5 cellspacing=0>
                     <TR bgcolor=Coral align=center>
                       <TD><B>SERVER NAME</B></TD>
                       <TD><B>RPC USERS</B></TD>
                       <TD><B>RPC CNX</B></TD>
                       <TD><B>OWA UNIQ USERS</B></TD>
                       <TD><B>EAS CUR REQ</B></TD>
                       <TD><B>EAS REQ/sec</B></TD>
                       <TD><B>IIS TOT ACTV REQ</B></TD>
                       <TD><B>IIS CUR CNX</B></TD>                  
                       <TD><B>IIS MAX CNX</B></TD></TR>"
                 
    Foreach($Entry in $Results) 
    
        {        
           $Outputreport += "<TR>"                      
           $Outputreport += "<TD>$($Entry.Servername)</TD>`
                            <TD align=center>$($Entry.RPCUser)</TD>`
                            <TD align=center>$($Entry.RPCCon)</TD>`
                            <TD align=center>$($Entry.OWA)</TD>`
                            <TD align=center>$($Entry.EAS1)</TD>`
                            <TD align=center>$($Entry.EAS2)</TD>`
                            <TD align=center>$($Entry.W3SVC)</TD>`
                            <TD align=center>$($Entry.WEB1)</TD>`
                            <TD align=center>$($Entry.WEB2)</TD></TR>" 
        }
     $Outputreport += "</Table></BODY></HTML>"

$Outputreport | out-file C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Connections_Report.htm
Invoke-Expression C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Connections_Report.htm

##Send email functionality from below line, use it if you want

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___Connections_Report.htm"

$SMTPServers = @("Server1","Server2","Server3")
#$SMTPServer1 = "EIS-LS2-EP06143"
#$SMTPServer2 = "EIS-LS2-EP06144"
#$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - Connections Report - $ReportDate"


if (Test-Connection $SMTPServers[0] -Quiet)
	{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $($SMTPServers[0]) -Attachments $EMailAttachment}
elseif (Test-Connection $SMTPServers[1] -Quiet)
	{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $($SMTPServers[1]) -Attachments $EMailAttachment}
else 
	{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $($SMTPServers[2]) -Attachments $EMailAttachment}