#################################################
#                                                                                                      
#    Exchange Mailbox Export Script                                        
#      Tested on Exchange 2016	                                            
#                                                                                                   
#    v.1.0 - Sebastian Storholm 13.05.2020                          
#                                                                                                    
#################################################

# Script exports the specified users mailbox as a PST to the specified path with the filename [alias]_YYYYMMDD.pst 

$UserCredential = Get-Credential
$SessionExchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.example.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $SessionExchange -DisableNameChecking

Write-Host "Mailbox PST Exporter v.1.0"

$Alias = Read-Host -Prompt 'User Alias?';
$user = $Alias.replace( ".","_")
$Date = get-date -f yyyyMMdd
$path = “\\SERVER1\PST Backups\Exported\" + $user + "_" + $Date + ".pst”
New-MailboxExportRequest -Mailbox $Alias -FilePath $path

# Check progress using either of these:
# Get-MailboxExportRequestStatistics -Identity $Alias
# Get-MailboxExportRequest | Get-MailboxExportRequestStatistics