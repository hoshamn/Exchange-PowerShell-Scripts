#############################################################################
#       Author: Vikas Sukhija
#       Date: 12/25/2014
#	Edited By: Hisham Nasur    
#       Date: 03/16/2022
#       Satus: (Ping and all Exchange Services)
#       Update: Added Advertising
#       Description: Exchange Services Health Status
#############################################################################
###########################Define Variables##################################

$reportpath = ".\ExchangeServiceStatus.htm" 

if((test-path $reportpath) -like $false)
{
new-item $reportpath -type file
}
$smtphost = "webmail.hniglabs.com"
$from = "ExchangeServiceCheck@hniglabs.com" 
$email1 = "administrator@hniglabs.com"
$timeout = "60"

###############################HTml Report Content############################
$report = $reportpath

Clear-Content $report 
Add-Content $report "<html>" 
Add-Content $report "<head>" 
Add-Content $report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $report '<title>Exchange Service Status Report</title>' 
add-content $report '<STYLE TYPE="text/css">' 
add-content $report  "<!--" 
add-content $report  "td {" 
add-content $report  "font-family: Tahoma;"
add-content $report  "font-size: 10px;" 
add-content $report  "border-top: 1px solid #999999;" 
add-content $report  "border-right: 1px solid #999999;" 
add-content $report  "border-bottom: 1px solid #999999;" 
add-content $report  "border-left: 1px solid #999999;" 
add-content $report  "padding-top: 3px;" 
add-content $report  "padding-right: 3px;" 
add-content $report  "padding-bottom: 3px;" 
add-content $report  "padding-left: 3px;" 
add-content $report  "}" 
add-content $report  "body {" 
add-content $report  "margin-left: 5px;" 
add-content $report  "margin-top: 5px;" 
add-content $report  "margin-right: 0px;" 
add-content $report  "margin-bottom: 10px;" 
add-content $report  "" 
add-content $report  "table {" 
add-content $report  "border: thin solid #000000;" 
add-content $report  "}" 
add-content $report  "-->" 
add-content $report  "</style>" 
Add-Content $report "</head>" 
Add-Content $report "<body>" 
add-content $report  "<table width='100%'>" 
add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='Auto' align='center'>" 
add-content $report  "<font face='tahoma' color='#00000' size='4'><strong>Exchange Service Health Status</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>" 
add-content $report  "</table>" 
 
add-content $report  "<table width='100%'>" 
Add-Content $report  "<tr bgcolor='Lavender'>"

Add-Content $report  "<td width='5%' align='center'><B>Identity</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>PingStatus</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeADTopology</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeAntispamUpdate</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeCompliance</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeDagMgmt</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeDelivery</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeDiagnostics</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeEdgeSync</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeFastSearch</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeFrontEndTransport</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeHM</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeHMRecovery</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeImap4</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeIMAP4BE</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeIS</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeMailboxAssistants</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeMailboxReplication</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeMitigation</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeNotificationsBroker</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangePop3</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangePOP3BE</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeRepl</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeRPC</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeServiceHost</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeSubmission</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeThrottling</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeTransport</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeTransportLogSearch</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeUM</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>MSExchangeUMCR</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>SearchExchangeTracing</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>vmickvpexchange</B></td>"
Add-Content $report  "<td width='10%' align='center'><B>wsbexchange</B></td>"
 
Add-Content $report "</tr>" 

#####################################Get ALL Exchange Servers#################################

$GetExchServers = Get-exchangeserver 
$ExchangeServers = Get-Service *Exchange* -ComputerName $GetExchServers


################Ping Test######

foreach ($exchange in $GetExchServers){
$Identity = $exchange
                Add-Content $report "<tr>"
if ( Test-Connection -ComputerName $exchange -Count 1 -ErrorAction SilentlyContinue ) {
Write-Host $exchange `t $exchange `t Ping Success -ForegroundColor Green
 
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
                Add-Content $report "<td bgcolor= '#34eb0a' align=center>  <B>Success</B></td>"

                ##############MSExchangeADTopology Service Status################
		$serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeADTopology" -ErrorAction SilentlyContinue
                
                
                 if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
                
               ######################################################
                ##############MSExchangeAntispamUpdate Service Status################
		$serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeAntispamUpdate" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
               ######################################################
                ##############MSExchangeCompliance Service Status################
		$serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeCompliance" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
               ######################################################

               ####################MSExchangeDagMgmt Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeDagMgmt" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
               ########################################################
               ####################MSExchangeDelivery Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeDelivery" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
               ########################################################
			####################MSExchangeDiagnostics Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeDiagnostics" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################	
			####################MSExchangeEdgeSync Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeEdgeSync" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
	       ####################MSExchangeFastSearch Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeFastSearch" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeFrontEndTransport Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeFrontEndTransport" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeHM Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeHM" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeHMRecovery Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeHMRecovery" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeImap4 Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeImap4" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeIMAP4BE Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeIMAP4BE" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeIS Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeIS" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeMailboxAssistants Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeMailboxAssistants" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeMailboxReplication Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeMailboxReplication" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeMitigation Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeMitigation" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeNotificationsBroker Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeNotificationsBroker" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangePop3 Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangePop3" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangePOP3BE Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangePOP3BE" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeRepl Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeRepl" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeRPC Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeRPC" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeServiceHost Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeServiceHost" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeSubmission Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeSubmission" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeThrottling Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeThrottling" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeTransport Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeTransport" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeTransportLogSearch Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeTransportLogSearch" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeUM Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeUM" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################MSExchangeUMCR Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "MSExchangeUMCR" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  }
			   ########################################################
		   ####################SearchExchangeTracing Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "SearchExchangeTracing" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################vmickvpexchange Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "vmickvpexchange" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
		   ####################wsbexchange Service status##################
               $serviceStatus = get-service -ComputerName $exchange -Name "wsbexchange" -ErrorAction SilentlyContinue
                if ($serviceStatus.status -like "Running") {
 		   Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Green 
         	   $svcName = $serviceStatus.name 
         	   $svcState = $serviceStatus.status          
         	   Add-Content $report "<td bgcolor= '#34eb0a' align=center><B>$svcState</B></td>" 
                  }
                 else 
                  { 
       		  Write-Host $exchange `t $serviceStatus.name `t $serviceStatus.status -ForegroundColor Red 
         	  $svcName = $serviceStatus.name 
         	  $svcState = $serviceStatus.status          
         	  Add-Content $report "<td bgcolor= 'Red' align=center><B>$svcState</B></td>" 
                  } 
			   ########################################################
                
} 
else
              {
Write-Host $exchange `t $exchange `t Ping Fail -ForegroundColor Red
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
                Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>" 
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>" 
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>" 
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>" 
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>"
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>"
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>"
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>"
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>Ping Fail</B></td>"
}         
       
} 

Add-Content $report "</tr>"
############################################Close HTMl Tables###########################


Add-content $report  "</table>" 
Add-Content $report "</body>" 
Add-Content $report "</html>" 


########################################################################################
#############################################Send Email#################################


$subject = "Exchange Service Health Monitor" 
$body = Get-Content ".\ExchangeServiceStatus.htm" 
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
$msg = New-Object System.Net.Mail.MailMessage 
$msg.To.Add($email1)
$msg.from = $from
$msg.subject = $subject
$msg.body = $body 
$msg.isBodyhtml = $true 
$smtp.send($msg) 

########################################################################################

########################################################################################
 
         	
		