################################################################################# 
## 
## Server Health Check 
## Created by Sravan Kumar S  
## Date : 3 Mar 2014 
## Version : 1.0 
## Email: sravankumar.s@outlook.com   
## This scripts check the server Avrg CPU and Memory utlization along with C drive 
## disk utilization and sends an email to the receipents included in the script
################################################################################ 

#$ServerListFile = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\Exchange_Servers_incl_2016.txt"
#$ServerList = Get-Content $ServerListFile -ErrorAction SilentlyContinue 
# Building search substring logic based on server names <change for your company/department>
$ServerNamesSubstrings = @("EIS-*-EP0*","IMG-*-EP*","DCS-DCB-EP*")
$Logic = ""
For ($i=0;$i -lt $ServerNamesSubstrings.Count;$i++){
    If ($i -ne $ServerNamesSubstrings.Count-1){
        $Logic += "(`$_.Name -Like $($ServerNamesSubstrings[$i])) -or "
    } Else { # If it's the last -or condition, we don't put -or at the end of the string.
        $Logic += "(`$_.Name -Like $($ServerNamesSubstrings[$i]))"
    }
}

#BEGIN To test logic only
$Logic
exit
#END To test logic only

$ServerList = Get-ExchangeServer | where {(($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")) -and (($_.ServerRole -like "*HubTransport*") -or ($_.ServerRole -like "Mailbox*"))} | sort name | Select Name
#$ServerList = Get-ExchangeServer | where {$_.name -like "EIS-LS2-EP06*"} | sort name | Select Name

$Result = @()
$ReportDate = Get-Date
$WarningEvents = 0
$ErrorEvents = 0

$Outputreport = "<HTML><TITLE> In-House Email Services - CPU Health Report </TITLE>
                 <BODY background-color:peachpuff>
                 <font color =""#99000"" face=""Microsoft Tai le"">
                 <H2> In-House Email Services - CPU Health Report - $ReportDate</H2></font>
                 <Table border=2 cellpadding=5 cellspacing=0>
                 <TR bgcolor=Aqua align=center>
                   <TD><B>Server Name</B></TD>
                   <TD><B>UpTime Hrs (<168)</B></TD>
                   <TD><B>Avg. CPU Util. % (>90%)</B></TD>
                   <TD><B>Memory Util. % (>99%)</B></TD>
                   <TD><B>C Drive Free Space % (<8%)</B></TD></TR>"

ForEach($ComputerName in $ServerList) 
{

$AVGProc = $null
$OS = $null
$OSInfo = $null
$timespan = $null
$uptime = $null
$vol = $null    

$AVGProc = Get-WmiObject -computername $computername.Name win32_processor | 
Measure-Object -property LoadPercentage -Average | Select Average

$OS = Get-WmiObject -Class win32_operatingsystem -computername $computername.Name |
Select-Object @{Name = "MemoryUsage"; Expression = {"{0:N2}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}

$OSInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computername.Name -ErrorAction CONTINUE

$timespan = $OSInfo.ConvertToDateTime($OSInfo.LocalDateTime) - $OSInfo.ConvertToDateTime($OSInfo.LastBootUpTime)
            [int]$uptime = "{0:00}" -f $timespan.TotalHours
                
$vol = Get-WmiObject -Class win32_Volume -ComputerName $computername.Name -Filter "DriveLetter = 'C:'" |
Select-object @{Name = "C_PercentFree"; Expression = {"{0:N2}" -f  (($_.FreeSpace / $_.Capacity)*100) } }
    
$Result += [PSCustomObject] @{ 
        
        ServerName = $computername.Name
        UpTime = "$uptime"
        CPULoad = "$($AVGProc.Average)"
        MemLoad = "$($OS.MemoryUsage)"
        CDrive = "$($vol.'C_PercentFree')"
         
    }}
               
        Foreach($Entry in $Result) 
    
        {  if(([decimal]$Entry.UpTime -le 168) -and ([decimal]$Entry.CPULoad -lt 90) -and ([decimal]$Entry.MemLoad -lt 99) -and ([decimal]$Entry.CDrive -gt 8))

            {$Outputreport += "<TR bgcolor=Yellow>"
             $WarningEvents ++}
              
           elseif(([decimal]$Entry.CPULoad -ge 90) -OR ([decimal]$Entry.MemLoad -ge 99) -OR ([decimal]$Entry.CDrive -le 8))
                
            {$Outputreport += "<TR bgcolor=Red>"
             $ErrorEvents ++}
            
           else
            {$Outputreport += "<TR>"                      }
             $Outputreport += "<TD>$($Entry.Servername)</TD><TD align=center>$($Entry.Uptime)</TD><TD align=center>$($Entry.CPULoad)</TD><TD align=center>$($Entry.MemLoad)</TD><TD align=center>$($Entry.Cdrive)</TD></TR>" 
            }
          
   $Outputreport += "</Table></BODY></HTML>"

$Outputreport | out-file C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___CPU_Health_Report.htm
#Invoke-Expression C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___CPU_Health_Report.htm

##Send email functionality from below line, use it if you want

if ($ErrorEvents -gt 0)
    {$EmailPriority = "High"}

elseif ($WarningEvents -gt 0)
    {$EmailPriority = "Normal"}
else
    {$EmailPriority = "Low"}

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\In-House_Email_Services___CPU_Health_Report.htm"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "kwok-fai.ha@ssc-spc.gc.ca"

$Subject = "In-House Email Services - CPU Health Report - $ReportDate"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -Priority $EmailPriority -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}