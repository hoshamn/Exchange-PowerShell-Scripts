(Get-DatabaseAvailabilityGroup -Identity (Get-MailboxServer -Identity $env:computername).DatabaseAvailabilityGroup).Servers | Test-MapiConnectivity | Sort Database | Format-Table -AutoSize
Get-MailboxDatabase | Sort Name | Get-MailboxDatabaseCopyStatus | Format-Table -AutoSize
function CopyCount 
{
$DatabaseList = Get-MailboxDatabase | Sort Name
$DatabaseList | % {
$Results = $_ | Get-MailboxDatabaseCopyStatus
$Good = $Results | where { ($_.Status -eq "Mounted") -or ($_.Status -eq "Healthy") }
$_ | add-member NoteProperty "CopiesTotal" $Results.Count
$_ | add-member NoteProperty "CopiesFailed" ($Results.Count-$Good.Count)
}
$DatabaseList | sort copiesfailed -Descending | ft name,copiesTotal,copiesFailed -AutoSize 
}
CopyCount