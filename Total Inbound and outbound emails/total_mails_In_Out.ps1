$From = "4/20/2022"
 $To = "4/24/2022"
    
 $intSent = 0
 $intRec = 0
    
 $Mailboxes = Get-Mailbox -ResultSize unlimited  | where {$_.RecipientTypeDetails -eq "UserMailbox"}
    
 foreach ($Mailbox in $Mailboxes){
 Get-TransportService | Get-MessageTrackingLog -Sender $Mailbox.PrimarySmtpAddress -ResultSize Unlimited -Start $From -End $To | ForEach {
     If ($_.EventId -eq "RECEIVE" -and $_.Source -eq "SMTP") {
         $intSent ++
     }
 }
    
 Get-TransportService | Get-MessageTrackingLog -Recipients $Mailbox.PrimarySmtpAddress -ResultSize Unlimited -Start $From -End $To | ForEach {
     If ($_.EventId -eq "DELIVER") {
         $intRec ++
     }
 }
 }
    
 Write-Host "`nResult:`n--------------------`n"
 Write-Host "Total Sent:"$intSent""
 Write-Host "Total Receive:"$intRec""
 Write-Host "`n--------------------`n"