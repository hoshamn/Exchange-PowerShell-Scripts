########################################################################################################################################
# Requirements Before running script
# 1. Active Directory RSAT Tools
#    https://gallery.technet.microsoft.com/Install-RSAT-for-Windows-75f5f92f 
# 2. Install Office 365 MFA Connector (run in Internet Explore) 
#    https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application 
# 3. User must be member of following Security Groups
#    Exchange Recipient Management
#    Exchange Unified Messaging
########################################################################################################################################

########################################################################################################################################
# Microsoft have an aggressive time out period
# Please have the username and delegate username ready before running the script and if asked please enter your password again. 
########################################################################################################################################
$doamin = "contoso.com" #change to your admin account domain
$date = get-date
$logfile = $null
$logfile += "########################################################"
$logfile += "`n# $date - $env:USERNAME@$doamin"
$logfile += "`n########################################################"

########################################################
# Ignore all errors
########################################################
$ErrorActionPreference= 'silentlycontinue'

########################################################
# Load Office 365 Module
########################################################
If ($CreateEXOPSSession -eq $null) {
    $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
    . "$CreateEXOPSSession\CreateExoPSSession.ps1"
    #Office 365
    Connect-EXOPSSession -UserPrincipalName "$env:USERNAME@$doamin" -ConnectionUri https://outlook.office365.com/powershell-liveid/
}

########################################################
# Load Active Directory Module
########################################################
If (!(Get-module ActiveDirectory)) {
    Write-Host "Loading Active Directory Modules" -foregroundcolor Yellow
    Import-Module ActiveDirectory
}

CLS

########################################################
# Custom Varablies 
########################################################
$Employee =  $null
$ADUser = $null
$UserPrincipalName = $null
$Mailbox = $null
$Manager = $null
$NewManager = $null
$Employee =  Read-Host 'Please Enter Employees Email Address'
$logfile += "`nEmployee Terminated - $Employee"
Write-Host "`n"

########################################################
# Check user indetity exists 
########################################################
$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainList = @($objForest.Domains | Select-Object Name)
$Domains = $DomainList | foreach {$_.Name}
Write-Host "~Searching for User~" -ForegroundColor Magenta
foreach($Domain in $Domains)
{
    Write-Host "Checking  - " -ForegroundColor DarkYellow -NoNewline
    ($check) = Get-ADUser -Server $Domain  -Filter "EmailAddress -Like '$Employee'" -Properties *
    if ($check) {
        Write-Host "FOUND  : $Employee exists on $domain" -ForegroundColor Green
        $ADUser = $check
        $logfile += "`nEmployee Location - $domain"
    } else {
        Write-Warning "$Employee does not exists on $domain"
    }
}
If ($ADUser -eq $null) {
    Write-Host "`n"
    Write-Host "User not found on any domain - please check spelling, or contact Tier 3" -ForegroundColor Red
    pause 
}

########################################################
# Remove the license from the user if already licensed 
########################################################
Write-Host "~Searching for O365 License~" -ForegroundColor Magenta
Write-Host "Checking  - " -ForegroundColor DarkYellow -NoNewline
$UserPrincipalName = $ADUser.UserPrincipalName
$isLicensed = $null
($isLicensed) = (get-MsolUser -UserPrincipalName $UserPrincipalName).licenses.AccountSkuId 
if ($isLicensed -eq "False") {
        Write-Host "IGNORE : $Employee does not have an active license" -ForegroundColor Yellow
    } else {
        $logfile += "`nLicenses Found: -"
        Write-Host "FOUND  : License has been removed from $UserPrincipalName" -ForegroundColor Green
        (get-MsolUser -UserPrincipalName $UserPrincipalName).licenses.AccountSkuId |
        foreach{
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $_
            $logfile += "`n$_"
        }   
    }


###################################################
# Remove Unified Messaging
###################################################
Write-Host "~Checking Unified Messaging~" -ForegroundColor Magenta
Write-Host "Checking  - " -ForegroundColor DarkYellow -NoNewline
$Mailbox = Get-Mailbox -Identity $ADUser.UserPrincipalName -ResultSize Unlimited
if ($Mailbox.UMEnabled -eq "True") {
    Write-Warning "Unified Messaging must be Disabled, Trying to Disable now."
    $UMMailbox = Get-UMMailbox $UserPrincipalName
    While ($UMMailbox.UMEnabled -eq "True") {
        Write-Host -NoNewline  "." -ForegroundColor DarkGray
        Disable-UMMailbox -Identity $UserPrincipalName | Out-Null
        Sleep 5
        $UMMailbox = Get-UMMailbox $UserPrincipalName
        }
        Write-Host "          - UPDATE : Unified Messaging has been Disabled" -ForegroundColor Green
        $logfile += "`nUnified Messaging Disabled - " + $UMMailbox.PhoneNumber
    } else {
        Write-Host "IGNORE : Unified Messaging Account not found on $Employee" -ForegroundColor Yellow
    }

###################################################
# Assigning access permissions to the Mailbox
###################################################
Write-Host "~Assigning Delegate Access to Mailbox~" -ForegroundColor Magenta
Write-Host "Assigning Permissions to $UserPrincipalName" -ForegroundColor DarkYellow
    $Delegate =  Read-Host 'Please Enter Delegate Email Address'
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $DomainList = @($objForest.Domains | Select-Object Name)
    $Domains = $DomainList | foreach {$_.Name}
    foreach($Domain in $Domains)
    {
        Write-Host "Checking  - " -ForegroundColor DarkYellow -NoNewline
        ($check)=Get-ADUser -Server $Domain  -Filter "EmailAddress -Like '$Delegate'" -Properties *
        if ($check) {
            Write-Host "FOUND  : $Delegate exists on $domain" -ForegroundColor Green
            $Manager = $check
        } else {
            Write-Warning "$Delegate does not exists on $domain"
        }
    }

    If ($Manager -eq $null) {
        Write-Warning "Delegate not found on any domain - Please manually add permissions using the O365 portal"
        $logfile += "`nPermission Assignment - No Manager Added"
    } else {
        Add-MailboxPermission -Identity $UserPrincipalName -User $Manager.Name -AccessRights FullAccess
        $logfile += "`nPermission Assignment - $Manager"
    }

    $MailboxPermissions = Get-MailboxPermission $UserPrincipalName
    if ($MailboxPermissions.User -contains $Manager.EmailAddress) {
        Write-Host "UPDATE - Delegate has been updated please instruct" $Manager.Name "how to access Delegate Mailbox" -ForegroundColor Green
    } else {
        Write-Host "UPDATE - Delegate has failed to update please esculate to Exchange team to set delegate access" -ForegroundColor red
    }

########################################################
# Hiden Account from GAL and OAB
########################################################
Set-Mailbox -Identity $ADUser.SamAccountName -HiddenFromAddressListsEnabled:$true

########################################################
# Update AD User Description 
########################################################
Write-Host "~Updating Description~" -ForegroundColor Magenta
Write-Host "Updating  - " -ForegroundColor DarkYellow -NoNewline
$Description = $ADUser.Description
$date = Get-Date (Get-Date).AddDays(90) -Format yyyy-MM-dd
Set-ADUser $ADUser.SamAccountName -Description "Disabled Ac - Delete $date - ($Description)"
Write-Host "UPDATED: Description has been updated" -ForegroundColor Green
$logfile += "`nDescription - 'Disabled Ac - Delete $date - ($Description)'"

########################################################
# Move AD Account to Accounts to be Deleted
########################################################
Write-Host "~Moving User Account~" -ForegroundColor Magenta
Write-Host "Moving    - " -ForegroundColor DarkYellow -NoNewline
$DistinguishedName = $ADUser.DistinguishedName
if ($DistinguishedName -match "DC=Contoso,DC=net") { #Change domain
    Write-Host "MOVED  : User is in Contoso domain, moving account to ACCOUNT TO BE DELETED" -ForegroundColor Green
    Get-ADUser $ADUser.SamAccountName | Move-ADObject -TargetPath "OU=ACCOUNTS TO BE DELETED,OU=User Accounts,DC=contoso,DC=net" # Change Location
    Disable-ADAccount -Identity $ADUser.SamAccountName
    $file = '\\contoso-file.net\Logs\termination.log' # Change location 
    $backup = '\\contoso-file.net\Logs\termination.bck'# Change location 
    
    if (Test-Path $file) { 
        if((Get-Item $file).length -gt 2048kb){
            Write-Host "Log File too Large - Creating new"
            Remove-Item -Path $backup
            Rename-Item -Path $file -NewName "termination.bck"
            New-Item $file -ItemType file
        }
    }
    $logfile += "`n"
    Add-content $file $logfile
}

Remove-Module ActiveDirectory
Get-PSSession | Remove-PSSession
$CreateEXOPSSession = $null