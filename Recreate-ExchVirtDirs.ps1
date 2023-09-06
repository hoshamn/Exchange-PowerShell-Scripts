add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

$Server = Read-Host "Enter Exchange Server name"
$Domain = Read-Host "Enter Root Domain (Domain.com)"
$BaseURL = Read-Host "Enter the Base URL (Mail.domain.com)"

Remove-ActiveSyncVirtualDirectory -Identity "$Server\Microsoft-Server-ActiveSync (Default Web Site)" -Confirm:$false -verbose
New-ActiveSyncVirtualDirectory -Server $server -InternalUrl "https://$BaseURL/Microsoft-Server-ActiveSync" -ExternalUrl "https://$BaseURL/Microsoft-Server-ActiveSync" -verbose

Remove-AutodiscoverVirtualDirectory -Identity "$Server\Autodiscover (Default Web Site)" -Confirm:$false -verbose
New-AutodiscoverVirtualDirectory -Server $Server -BasicAuthentication $true -WindowsAuthentication $true -verbose
Set-ClientAccessServer -Identity $Server -AutodiscoverServiceInternalUri https://autodiscover.$Domain/Autodiscover/Autodiscover.xml -verbose

Remove-EcpVirtualDirectory -Identity "$Server\ecp (Default Web Site)" -Confirm:$false -verbose
New-EcpVirtualDirectory -Server $Server -InternalUrl "https://$BaseURL/ecp" -verbose

Remove-MapiVirtualDirectory -Identity "$Server\mapi (Default Web Site)" -Confirm:$false -verbose
New-MapiVirtualDirectory -Server $Server -InternalUrl https://$BaseURL/mapi -IISAuthenticationMethods Ntlm, OAuth, Negotiate -verbose

Remove-OabVirtualDirectory -Identity "$Server\OAB (Default Web Site)" -Confirm:$false -Force -verbose
New-OabVirtualDirectory -Server $Server -InternalUrl "https://$BaseURL/OAB" -verbose

Remove-OwaVirtualDirectory -Identity "$Server\owa (Default Web Site)" -Confirm:$false -verbose
New-OwaVirtualDirectory -Server $Server -InternalUrl "https://$BaseURL/owa" -ExternalUrl "https://$BaseURL/owa" -verbose

Remove-PowerShellVirtualDirectory -Identity "$Server\PowerShell (Default Web Site)" -Confirm:$false -verbose
New-PowerShellVirtualDirectory -Server $Server -Name Powershell -InternalUrl https://$BaseURL/PowerShell -ExternalUrl https://$BaseURL/PowerShell -BasicAuthentication $false -WindowsAuthentication $true -CertificateAuthentication $false -verbose

Remove-WebServicesVirtualDirectory -Identity "$Server\EWS (Default Web Site)" -Confirm:$false -Force -verbose
New-WebServicesVirtualDirectory -Server $Server -InternalUrl "https://$BaseURL/EWS/Exchange.asmx" -ExternalUrl "https://$BaseURL/EWS/Exchange.asmx" -verbose

IISRESET