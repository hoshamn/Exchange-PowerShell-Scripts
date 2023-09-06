$file = 'C:\NourNet\ScheduleTasks\Exchange-Report\Exchange_Mailbox.csv'

Get-Mailbox -ResultSize unlimited -Filter * | select DisplayName, Alias, RecipientTypeDetails, PrimarySmtpAddress, WhenCreatedUTC, WhenMailboxCreated, WasInactiveMailbox, AccountDisabled, MaxSendSize, MaxReceiveSize, IssueWarningQuota, ProhibitSendReceiveQuota, HiddenFromAddressListsEnabled, OrganizationalUnit, ServerName, Database, @{label="TotalItemSize(MB)";expression={(get-mailboxstatistics $_).TotalItemSize.Value.ToMB()}}, @{label="ItemCount";expression={(get-mailboxstatistics $_).ItemCount}} | Export-Csv -Force -NoTypeInformation -Path $file 

$options = @{
    'SmtpServer' = "10.214.15.142" 
    'To' = "m.kotb@nour.net.sa","samirah@nour.net.sa","ms-support@nour.net.sa"
    'From' = "Exch_Report@spga.gov.sa"
    'Subject' = "SPGA Exchange mailbox monthly Report" 
    'Body' = "Please find attached spreadsheet contains SPFA Exchange mailbox monthly Report" 
    'Attachments' = $file  
}

Send-MailMessage @options
