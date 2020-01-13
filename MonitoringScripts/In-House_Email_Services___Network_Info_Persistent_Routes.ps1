#########################################################################
#			Author: Diwakar Sharma
#			Reviewer: Vikas Sukhija
#			Date: 06/10/2015
#			Reviewed: 06/15/2015
#			Desc: Collect Network info from Nics
#
# Last updated:  KH 2019 03 01
#
#           Modified to gather network static routes      
#########################################################################

$Collection = @()

#$ComputerNAme = get-content "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\NITL_DEMS_Exchange_Servers_incl_2016.txt"
$ComputerName = Get-ExchangeServer | where {(($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")) -and (($_.ServerRole -like "*HubTransport*") -or ($_.ServerRole -like "Mailbox*"))} | sort name | Select Name

foreach ($Computer in $ComputerName) {
#$COMPUTER
  if(Test-Connection -ComputerName $Computer.Name -Count 1 -ea 0) {
   
   $IP4PersistedRouteTable = $null
   $IP4PersistedRouteTable = Get-WmiObject Win32_IP4PersistedRouteTable -ComputerName $Computer.Name
        
if($IP4PersistedRouteTable){
    
foreach ($IP4PersistedRoute in $IP4PersistedRouteTable) {

    $__SERVER = $null
    $Name = $null
    $Destination = $null
    $Mask = $null
    $NextHop = $null
    $Metric1 = $null
    
    $__SERVER = $IP4PersistedRoute.__SERVER
    $Name = $IP4PersistedRoute.Name
    $Destination = $IP4PersistedRoute.Destination
    $Mask = $IP4PersistedRoute.Mask
    $NextHop = $IP4PersistedRoute.NextHop
    $Metric1 = $IP4PersistedRoute.Metric1
        	        
    $OutputObj = New-Object -Type PSObject

    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $__SERVER.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name Name -Value $Name
    $OutputObj | Add-Member -MemberType NoteProperty -Name Network_Address -Value $Destination
    $OutputObj | Add-Member -MemberType NoteProperty -Name Netmask -Value $Mask
    $OutputObj | Add-Member -MemberType NoteProperty -Name Gateway_Address -Value $NextHop
    $OutputObj | Add-Member -MemberType NoteProperty -Name Metric -Value $Metric1
        
    $OutputObj

$Collection += $OutputObj

          }
      }
 }

}

#Create variable for log date

$LogDate = Get-Date -f yyyy-MM-dd_HH-mm-ss

#Export report to CSV file 

$Collection | Export-Csv -path "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\Network_Info\In-House_Email_Services___Network_Info_Persistent_Routes_$logDate.csv" -NoTypeInformation

#Send disk report using the exchange email module

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\Network_Info\In-House_Email_Services___Network_Info_Persistent_Routes_$logDate.csv"

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "Kwok-Fai.HA@tdc.forces.gc.ca"

$Subject = "In-House Email Services - Network Info Persistent Routes"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}

###############################################################################