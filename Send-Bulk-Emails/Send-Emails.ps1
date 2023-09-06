
# Function to create report email
function SendNotification{
$Msg = New-Object Net.Mail.MailMessage
$Smtp = New-Object Net.Mail.SmtpClient($ExchangeServer)
$Msg.From = $FromAddress
$Msg.To.Add($ToAddress)
$Msg.Subject = "Announcement: Important information about your IT Department."
$Msg.Body = $EmailBody
$Msg.IsBodyHTML = $true
$Smtp.Send($Msg)
}

# Define local Exchange server info for message relay. Ensure that any servers running this script have permission to relay.
$ExchangeServer = "EX19-MBX01.hniglabs.local"
$FromAddress = "administrator@hniglabs.com"

# Import user list and information from .CSV file
$Users = Import-Csv UserList.csv

# Send notification to each user in the list
Foreach ($User in $Users) {
$ToAddress = $User.Email
$Name = $User.FirstName
$Level = $User.Level
$DeskNum = $User.DeskNumber
$PhoneNum = $User.PhoneNumber
$EmailBody = @"



Dear $Name,

As you know we will be relocating to our new offices at 742 Evergreen Terrace, Springfield on July 1, 2015. This email contains important information to help you get settled as quickly as possible.

Your existing access card will grant you access to the new building and your desk location is as follows:

Your Email Address is : $ToAddress
Desk Number: $DeskNum
Phone Number: $PhoneNum

Your new phone will be connected and ready for use when you arrive.

If you require any assistance during the move please contact the relocation helpdesk at Helpdesk@hniglabs.com or by calling 555-555-1234

Regards,

IT Team



"@
Write-Host "Sending notification to $Name ($ToAddress)" -ForegroundColor Yellow
SendNotification
}


