#########################################################################
#			Author: Diwakar Sharma
#			Reviewer: Vikas Sukhija
#			Date: 06/10/2015
#			Reviewed: 06/15/2015
#			Desc: Collect Network info from Nics
#
# Last updated:  KH 2019 03 01
#
#           Added many fields, WMI classes, e-mail capability        
#########################################################################

$Collection = @()

#$ComputerName = get-content "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\NITL_DEMS_Exchange_Servers_incl_2016.txt"
$ComputerName = Get-ExchangeServer | where {(($_.name -like "EIS-*-EP0*") -or ($_.name -like "IMG-*-EP*") -or ($_.name -like "DCS-DCB-EP*")) -and (($_.ServerRole -like "*HubTransport*") -or ($_.ServerRole -like "Mailbox*"))} | sort name | Select Name

foreach ($Computer in $ComputerName) {
#$COMPUTER
  if(Test-Connection -ComputerName $Computer.Name -Count 1 -ea 0) {
   $Networks = $null
   $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer.Name -ea silentlycontinue | ? {$_.IPEnabled}
        
if($Networks){
    
foreach ($Network in $Networks) {

    $IPAddress = $null
    $SubnetMask = $null
    $DefaultGateway= $null
    $DNSServers = $null
    $WINSPrimaryserver = $null
    $WINSSecondaryserver = $null
    $IsDHCPEnabled = $null

	$Description = $null
	$DNSDomain = $null
	$DNSDomainSuffixSearchOrder = $null
	$DNSEnabledForWINSResolution = $null
	$DomainDNSRegistrationEnabled = $null
	$FullDNSRegistrationEnabled = $null
	$IPFilterSecurityEnabled = $null
	$WINSEnableLMHostsLookup = $null
	$WINSHostLookupFile = $null
	$KeepAliveInterval = $null
    $KeepAliveTime = $null
  	$MACAddress = $null
  	$MTU = $null

    $NetConnectionID = $null
    $Speed = $null
	
    $IPAddress  = $Network.IpAddress[0]

    $SubnetMask  = $Network.IPSubnet[0]

    $DefaultGateway = $Network.DefaultIPGateway -join ','

    $DNSServers  = $Network.DNSServerSearchOrder -join ','

    $WINSPrimaryserver = $Networks.WINSPrimaryServer
    $WINSSecondaryserver = $Networks.WINSSecondaryserver
    
    $Description = $Network.Description
	$DNSDomain = $Network.DNSDomain
	$DNSDomainSuffixSearchOrder = $Network.DNSDomainSuffixSearchOrder -join ','
	$DNSEnabledForWINSResolution = $Network.DNSEnabledForWINSResolution
	$DomainDNSRegistrationEnabled = $Network.DomainDNSRegistrationEnabled
	$FullDNSRegistrationEnabled = $Network.FullDNSRegistrationEnabled
	$IPFilterSecurityEnabled = $Network.IPFilterSecurityEnabled
	$WINSEnableLMHostsLookup = $Network.WINSEnableLMHostsLookup
	$WINSHostLookupFile = $Network.WINSHostLookupFile
    $KeepAliveInterval = $Network.KeepAliveInterval
	$KeepAliveTime = $Network.KeepAliveTime
  	$MACAddress = $Network.MACAddress
  	$MTU = $Network.MTU
 
    $NetAdapterInfo = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $Computer.Name | ? {$_.MACAddress -like $MACAddress}

    $NetConnectionID = $NetAdapterInfo.NetConnectionID
    $Speed = $NetAdapterInfo.Speed

    $IsDHCPEnabled = $false

    If($network.DHCPEnabled) {
     $IsDHCPEnabled = $true
    }

    $OutputObj  = New-Object -Type PSObject
    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.Name.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name Network_Connection_Name -Value $NetConnectionID

    $OutputObj | Add-Member -MemberType NoteProperty -Name Connect_using -Value $Description
    $OutputObj | Add-Member -MemberType NoteProperty -Name Obtain_an_IP_address_automatically -Value $IsDHCPEnabled

    $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name SubnetMask -Value $SubnetMask
    $OutputObj | Add-Member -MemberType NoteProperty -Name Gateway -Value $DefaultGateway
        
    $OutputObj | Add-Member -MemberType NoteProperty -Name DNSServers -Value $DNSServers
    $OutputObj | Add-Member -MemberType NoteProperty -Name DNS_suffixes -Value $DNSDomainSuffixSearchOrder

    $OutputObj | Add-Member -MemberType NoteProperty -Name DNS_suffix_for_this_connection -Value $DNSDomain
    $OutputObj | Add-Member -MemberType NoteProperty -Name Register_this_connection"'"s_addresses_in_DNS -Value $FullDNSRegistrationEnabled
    $OutputObj | Add-Member -MemberType NoteProperty -Name Use_this_connection"'"s_DNS_suffix_in_DNS_registration -Value $DomainDNSRegistrationEnabled

    $OutputObj | Add-Member -MemberType NoteProperty -Name WINSPrimaryserver -Value $WINSPrimaryserver
    $OutputObj | Add-Member -MemberType NoteProperty -Name WINSSecondaryserver -Value $WINSSecondaryserver

    $OutputObj | Add-Member -MemberType NoteProperty -Name Enable_LMHOSTS_lookup -Value $WINSEnableLMHostsLookup
    $OutputObj | Add-Member -MemberType NoteProperty -Name WINSHostLookupFile -Value $WINSHostLookupFile
    $OutputObj | Add-Member -MemberType NoteProperty -Name DNSEnabledForWINSResolution -Value $DNSEnabledForWINSResolution

    $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name IPFilterSecurityEnabled -Value $IPFilterSecurityEnabled
    
    $OutputObj | Add-Member -MemberType NoteProperty -Name KeepAliveInterval -Value $KeepAliveInterval
    $OutputObj | Add-Member -MemberType NoteProperty -Name KeepAliveTime -Value $KeepAliveTime
    
    $OutputObj | Add-Member -MemberType NoteProperty -Name Speed -Value $Speed
    $OutputObj | Add-Member -MemberType NoteProperty -Name MTU -Value $MTU

    $OutputObj

$Collection += $OutputObj

          }
      }
 }

}

#Create variable for log date

$LogDate = Get-Date -f yyyy-MM-dd_HH-mm-ss

#Export report to CSV file 

$Collection | Export-Csv -path "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\Network_Info\In-House_Email_Services___Network_Info_$logDate.csv" -NoTypeInformation

#Send disk report using the exchange email module

$EMailAttachment = "C:\Scripts\Daily_Checks\Scheduled_Tasks-EMailed_Results\Network_Info\In-House_Email_Services___Network_Info_$logDate.csv"

#send-mailmessage -from Admin@forces.gc.ca -to DEMSCEMOPS@forces.gc.ca -subject "NITL DEMS Network Info" -SMTPServer IMG-LSL-EP00895 -Attachments $EMailAttachment

$SMTPServer1 = "EIS-LS2-EP06143"
$SMTPServer2 = "EIS-LS2-EP06144"
$SMTPServer3 = "EIS-DCB-EP06157"

$From = "Admin@forces.gc.ca"
$Recipient = "DEMSCEMOPS@forces.gc.ca"
#$Recipient = "Kwok-Fai.HA@tdc.forces.gc.ca"

$Subject = "In-House Email Services - Network Info"

if (Test-Connection $SMTPServer1 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer1 -Attachments $EMailAttachment}

elseif (Test-Connection $SMTPServer2 -Quiet)

{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer2 -Attachments $EMailAttachment}

else 
 
{Send-MailMessage -from $From -to $Recipient -subject $Subject -SMTPServer $SMTPServer3 -Attachments $EMailAttachment}

###############################################################################