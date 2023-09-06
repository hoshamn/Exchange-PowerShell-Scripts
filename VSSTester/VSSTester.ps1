<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 21.08.18.1629

#################################################################################
# Purpose:
# This script will allow you to test VSS functionality on Exchange server using DiskShadow.
# The script will automatically detect active and passive database copies running on the server.
# The general logic is:
# - start a PowerShell transcript
# - enable ExTRA tracing
# - enable VSS tracing
# - optionally: create the diskshadow config file with shadow expose enabled,
#               execute VSS backup using diskshadow,
#               delete the VSS snapshot post-backup
# - stop PowerShell transcript
#
#################################################################################
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'Allowing empty catch blocks for now as we need to be able to handle the exceptions.')]
[CmdletBinding()]
param(
)



Function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [scriptblock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

Function Invoke-CatchActionErrorLoop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]$CurrentErrors,
        [Parameter(Mandatory = $false, Position = 1)]
        [scriptblock]$CatchActionFunction
    )
    process {
        if ($null -ne $CatchActionFunction -and
            $Error.Count -ne $CurrentErrors) {
            $i = 0
            while ($i -lt ($Error.Count - $currentErrors)) {
                & $CatchActionFunction $Error[$i]
                $i++
            }
        }
    }
}

Function Confirm-ExchangeShell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [bool]$LoadExchangeShell = $true,

        [Parameter(Mandatory = $false)]
        [scriptblock]$CatchActionFunction
    )

    begin {
        $passed = $false
        $edgeTransportKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'
        $setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed: LoadExchangeShell: $LoadExchangeShell | Identity: $Identity"
        $params = @{
            Identity    = $Identity
            ErrorAction = "Stop"
        }
    }
    process {
        try {
            $currentErrors = $Error.Count
            Get-ExchangeServer @params | Out-Null
            Write-Verbose "Exchange PowerShell Module already loaded."
            $passed = $true
            Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
        } catch {
            Write-Verbose "Failed to run Get-ExchangeServer"
            Invoke-CatchActionError $CatchActionFunction

            if (-not ($LoadExchangeShell)) {
                return
            }

            #Test 32 bit process, as we can't see the registry if that is the case.
            if (-not ([System.Environment]::Is64BitProcess)) {
                Write-Warning "Open a 64 bit PowerShell process to continue"
                return
            }

            if (Test-Path "$setupKey") {
                $currentErrors = $Error.Count
                Write-Verbose "We are on Exchange 2013 or newer"

                try {
                    if (Test-Path $edgeTransportKey) {
                        Write-Verbose "We are on Exchange Edge Transport Server"
                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop

                        foreach ($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn) {
                            Write-Verbose "Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name
                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                        }

                        Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop
                    } else {
                        Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                    }

                    Write-Verbose "Imported Module. Trying Get-Exchange Server Again"
                    Get-ExchangeServer @params | Out-Null
                    $passed = $true
                    Write-Verbose "Successfully loaded Exchange Management Shell"
                    Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
                } catch {
                    Write-Warning "Failed to Load Exchange PowerShell Module..."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } else {
                Write-Verbose "Not on an Exchange or Tools server"
            }
        }
    }
    end {

        $currentErrors = $Error.Count
        $returnObject = [PSCustomObject]@{
            ShellLoaded = $passed
            Major       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMajor" -ErrorAction SilentlyContinue).MsiProductMajor)
            Minor       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMinor" -ErrorAction SilentlyContinue).MsiProductMinor)
            Build       = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMajor" -ErrorAction SilentlyContinue).MsiBuildMajor)
            Revision    = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMinor" -ErrorAction SilentlyContinue).MsiBuildMinor)
            EdgeServer  = $passed -and (Test-Path $setupKey) -and (Test-Path $edgeTransportKey)
            ToolsOnly   = $passed -and (Test-Path $setupKey) -and (!(Test-Path $edgeTransportKey)) -and `
            ($null -eq (Get-ItemProperty -Path $setupKey -Name "Services" -ErrorAction SilentlyContinue))
            RemoteShell = $passed -and (!(Test-Path $setupKey))
        }

        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

        return $returnObject
    }
}

Function Write-HostWriter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Need to use Write Host')]
    param(
        [Parameter(Mandatory = $true)][string]$WriteString
    )
    if ($null -ne $Script:Logger) {
        $Script:Logger.WriteHost($WriteString)
    } elseif ($null -eq $HostFunctionCaller) {
        Write-Host $WriteString
    } else {
        &$HostFunctionCaller $WriteString
    }
}

Function Write-VerboseWriter {
    param(
        [Parameter(Mandatory = $true)][string]$WriteString
    )
    if ($null -ne $Script:Logger) {
        $Script:Logger.WriteVerbose($WriteString)
    } elseif ($null -eq $VerboseFunctionCaller) {
        Write-Verbose $WriteString
    } else {
        &$VerboseFunctionCaller $WriteString
    }
}

function Invoke-CreateDiskShadowFile {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWMICmdlet', '', Justification = 'Required to get drives on old systems')]
    param()

    function Out-DHSFile {
        param ([string]$fileline)
        $fileline | Out-File -FilePath "$path\diskshadow.dsh" -Encoding ASCII -Append
    }

    #	creates the diskshadow.dsh file that will be written to below
    #	-------------------------------------------------------------
    $nl
    Get-Date
    Write-Host "Creating diskshadow config file..." -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    $nl
    New-Item -Path $path\diskshadow.dsh -type file -Force | Out-Null

    #	beginning lines of file
    #	-----------------------
    Out-DHSFile "set verbose on"
    Out-DHSFile "set context persistent"
    Out-DHSFile " "

    #	writers to exclude
    #	------------------
    Out-DHSFile "writer exclude {e8132975-6f93-4464-a53e-1050253ae220}"
    Out-DHSFile "writer exclude {2a40fd15-dfca-4aa8-a654-1f8c654603f6}"
    Out-DHSFile "writer exclude {35E81631-13E1-48DB-97FC-D5BC721BB18A}"
    Out-DHSFile "writer exclude {be000cbe-11fe-4426-9c58-531aa6355fc4}"
    Out-DHSFile "writer exclude {4969d978-be47-48b0-b100-f328f07ac1e0}"
    Out-DHSFile "writer exclude {a6ad56c2-b509-4e6c-bb19-49d8f43532f0}"
    Out-DHSFile "writer exclude {afbab4a2-367d-4d15-a586-71dbb18f8485}"
    Out-DHSFile "writer exclude {59b1f0cf-90ef-465f-9609-6ca8b2938366}"
    Out-DHSFile "writer exclude {542da469-d3e1-473c-9f4f-7847f01fc64f}"
    Out-DHSFile "writer exclude {4dc3bdd4-ab48-4d07-adb0-3bee2926fd7f}"
    Out-DHSFile "writer exclude {41e12264-35d8-479b-8e5c-9b23d1dad37e}"
    Out-DHSFile "writer exclude {12ce4370-5bb7-4C58-a76a-e5d5097e3674}"
    Out-DHSFile "writer exclude {cd3f2362-8bef-46c7-9181-d62844cdc062}"
    Out-DHSFile "writer exclude {dd846aaa-A1B6-42A8-AAF8-03DCB6114BFD}"
    Out-DHSFile "writer exclude {B2014C9E-8711-4C5C-A5A9-3CF384484757}"
    Out-DHSFile "writer exclude {BE9AC81E-3619-421F-920F-4C6FEA9E93AD}"
    Out-DHSFile "writer exclude {F08C1483-8407-4A26-8C26-6C267A629741}"
    Out-DHSFile "writer exclude {6F5B15B5-DA24-4D88-B737-63063E3A1F86}"
    Out-DHSFile "writer exclude {368753EC-572E-4FC7-B4B9-CCD9BDC624CB}"
    Out-DHSFile "writer exclude {5382579C-98DF-47A7-AC6C-98A6D7106E09}"
    Out-DHSFile "writer exclude {d61d61c8-d73a-4eee-8cdd-f6f9786b7124}"
    Out-DHSFile "writer exclude {75dfb225-e2e4-4d39-9ac9-ffaff65ddf06}"
    Out-DHSFile "writer exclude {0bada1de-01a9-4625-8278-69e735f39dd2}"
    Out-DHSFile "writer exclude {a65faa63-5ea8-4ebc-9dbd-a0c4db26912a}"
    Out-DHSFile " "

    #	add databases to exclude
    #	------------------------
    foreach ($db in $databases) {
        $dbg = ($db.guid)

        if (($db).guid -ne $dbGuid) {
            if (($db.IsMailboxDatabase) -eq "True") {
                $mountedOnServer = (Get-MailboxDatabase $db).server.name
            } else {
                $mountedOnServer = (Get-PublicFolderDatabase $db).server.name
            }
            if ($mountedOnServer -eq $serverName) {
                $script:activeNode = $true

                Out-DHSFile "writer exclude `"Microsoft Exchange Writer:\Microsoft Exchange Server\Microsoft Information Store\$serverName\$dbg`""
            }
            #if passive copy, add it with replica in the string
            else {
                $script:activeNode = $false
                Out-DHSFile "writer exclude `"Microsoft Exchange Replica Writer:\Microsoft Exchange Server\Microsoft Information Store\Replica\$serverName\$dbg`""
            }
        }
        #	add database to include
        #	-----------------------
        else {
            if (($db.IsMailboxDatabase) -eq "True") {
                $mountedOnServer = (Get-MailboxDatabase $db).server.name
            } else {
                $mountedOnServer = (Get-PublicFolderDatabase $db).server.name
            }
        }
    }
    Out-DHSFile " "
    Out-DHSFile "Begin backup"

    #	add the volumes for the included database
    #	-----------------------------------------
    #gets a list of mount points on local server
    $mpvolumes = Get-WmiObject -Query "select name, deviceid from win32_volume where drivetype=3 AND driveletter=NULL"
    $deviceIDs = @()

    #if selected database is a mailbox database, get mailbox paths
    if ((($databases[$dbtoBackup]).IsMailboxDatabase) -eq "True") {
        $getDB = (Get-MailboxDatabase $selDB)

        $dbMP = $false
        $logMP = $false

        #if no mountpoints ($mpvolumes) causes null-valued error, need to handle
        if ($null -ne $mpvolumes) {
            foreach ($mp in $mpvolumes) {
                $mpname = (($mp.name).substring(0, $mp.name.length - 1))
                #if following mount point path exists in database path use deviceID in diskshadow config file
                if ($getDB.EdbFilePath.pathname.ToString().ToLower().StartsWith($mpname.ToString().ToLower())) {
                    Write-Host " "
                    Write-Host "Mount point:  $($mp.name) in use for database path: "
                    #Write-host "Yes. I am a database in mountpoint"
                    "The selected database path is: " + $getDB.EdbFilePath.pathname
                    Write-Host "adding deviceID to file: "
                    $dbEdbVol = $mp.deviceid
                    Write-Host $dbEdbVol

                    #add device ID to array
                    $deviceID1 = $mp.DeviceID
                    $dbMP = $true
                }

                #if following mount point path exists in log path use deviceID in diskshadow config file
                if ($getDB.LogFolderPath.pathname.ToString().ToLower().contains($mpname.ToString().ToLower())) {
                    Write-Host " "
                    Write-Host "Mount point: $($mp.name) in use for log path: "
                    #Write-host "Yes. My logs are in a mountpoint"
                    "The log folder path of selected database is: " + $getDB.LogFolderPath.pathname
                    Write-Host "adding deviceID to file: "
                    $dbLogVol = $mp.deviceid
                    Write-Host $dbLogVol
                    $deviceID2 = $mp.DeviceID
                    $logMP = $true
                }
            }
            $deviceIDs = $deviceID1, $deviceID2
        }
    }

    #if not a mailbox database, assume its a public folder database, get public folder paths
    if ((($databases[$dbtoBackup]).IsPublicFolderDatabase) -eq "True") {
        $getDB = (Get-PublicFolderDatabase $selDB)

        $dbMP = $false
        $logMP = $false

        if ($null -ne $mpvolumes) {
            foreach ($mp in $mpvolumes) {
                $mpname = (($mp.name).substring(0, $mp.name.length - 1))
                #if following mount point path exists in database path use deviceID in diskshadow config file

                if ($getDB.EdbFilePath.pathname.ToString().ToLower().StartsWith($mpname.ToString().ToLower())) {
                    Write-Host " "
                    Write-Host "Mount point: $($mp.name) in use for database path: "
                    "The current database path is: " + $getDB.EdbFilePath.pathname
                    Write-Host "adding deviceID to file: "
                    $dbEdbVol = $mp.deviceId
                    Write-Host $dbvol

                    #add device ID to array
                    $deviceID1 = $mp.DeviceID
                    $dbMP = $true
                }

                #if following mount point path exists in log path use deviceID in diskshadow config file
                if ($getDB.LogFolderPath.pathname.ToString().ToLower().contains($mpname.ToString().ToLower())) {
                    Write-Host " "
                    Write-Host "Mount point: $($vol.name) in use for log path: "
                    "The log folder path of selected database is: " + $getDB.LogFolderPath.pathname
                    Write-Host "adding deviceID to file "
                    $dbLogVol = $mp.deviceId
                    Write-Host $dblogvol

                    $deviceID2 = $mp.DeviceID
                    $logMP = $true
                }
            }
            $deviceIDs = $deviceID1, $deviceID2
        }
    }

    if ($dbMP -eq $false) {

        $dbEdbVol = ($getDB.EdbFilePath.pathname).substring(0, 2)
        "The selected database path is '" + $getDB.EdbFilePath.pathname + "' so adding volume $dbEdbVol to backup scope"
        $deviceID1 = $dbEdbVol
    }

    if ($logMP -eq $false) {
        $dbLogVol = ($getDB.LogFolderPath.pathname).substring(0, 2)
        $nl
        "The selected database log folder path is '" + $getDB.LogFolderPath.pathname + "' so adding volume $dbLogVol to backup scope"
        $deviceID2 = $dbLogVol
    }

    # Here is where we start adding the appropriate volumes or mountpoints to the diskshadow config file
    # We make sure that we add only one Logical volume when we detect the EDB and log files
    # are on the same volume

    $nl
    $deviceIDs = $deviceID1, $deviceID2
    $comp = [string]::Compare($deviceID1, $deviceID2, $True)
    if ($comp -eq 0) {
        $dID = $deviceIDs[0]
        Write-Debug -Message ('$dID = ' + $dID.ToString())
        Write-Debug "When the database and log files are on the same volume, we add the volume only once"
        if ($dID.length -gt "2") {
            $addVol = "add volume $dID alias vss_test_" + ($dID).tostring().substring(11, 8)
            Write-Host $addVol
            Out-DHSFile $addVol
        } else {
            $addVol = "add volume $dID alias vss_test_" + ($dID).tostring().substring(0, 1)
            Write-Host $addVol
            Out-DHSFile $addVol
        }
    } else {
        Write-Host " "
        foreach ($device in $deviceIDs) {
            if ($device.length -gt "2") {
                Write-Host "Adding the Mount Point for DSH file"
                $addVol = "add volume $device alias vss_test_" + ($device).tostring().substring(11, 8)
                Write-Host $addVol
                Out-DHSFile $addVol
            } else {
                Write-Host "Adding the volume for DSH file"
                $addVol = "add volume $device alias vss_test_" + ($device).tostring().substring(0, 1)
                Write-Host $addVol
                Out-DHSFile $addVol
            }
        }
    }
    Out-DHSFile "create"
    Out-DHSFile " "
    $nl
    Get-Date
    Write-Host "Getting drive letters for exposing backup snapshot" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------------------------------------------------------"

    # check to see if the drives are the same for both database and logs
    # if the same volume is used, only one drive letter is needed for exposure
    # if two volumes are used, two drive letters are needed

    $matchCondition = "^[a-z]:$"
    Write-Debug $matchCondition

    if ($dbEdbVol -eq $dbLogVol) {
        $nl
        "Since the same volume is used for this database's EDB and logs, we only need a single drive"
        "letter to expose the backup snapshot."
        $nl

        do {
            Write-Host "Enter an unused drive letter with colon (e.g. X:) to expose the snapshot" -ForegroundColor Yellow -NoNewline
            $script:dbsnapvol = Read-Host " "
            if ($dbsnapvol -notmatch $matchCondition) {
                Write-Host "Your input was not acceptable. Please use a single letter and colon, e.g. X:" -ForegroundColor red
            }
        } while ($dbsnapvol -notmatch $matchCondition)
    } else {
        $nl
        "Since different volumes are used for this database's EDB and logs, we need two drive"
        "letters to expose the backup snapshot."
        $nl

        do {
            Write-Host "Enter an unused drive letter with colon (e.g. X:) to expose the DATABASE volume" -ForegroundColor Yellow -NoNewline
            $script:dbsnapvol = Read-Host " "
            if ($dbsnapvol -notmatch $matchCondition) {
                Write-Host "Your input was not acceptable. Please use a single letter and colon, e.g. X:" -ForegroundColor red
            }
        } while ($dbsnapvol -notmatch $matchCondition)

        do {
            Write-Host "Enter an unused drive letter with colon (e.g. Y:) to expose the LOG volume" -ForegroundColor Yellow -NoNewline
            $script:logsnapvol = Read-Host " "
            if ($logsnapvol -notmatch $matchCondition) {
                Write-Host "Your input was not acceptable. Please use a single letter and colon, e.g. Y:" -ForegroundColor red
            }
            if ($logsnapvol -eq $dbsnapvol) {
                Write-Host "You must choose a different drive letter than the one chosen to expose the DATABASE volume." -ForegroundColor red
            }
        } while (($logsnapvol -notmatch $matchCondition) -or ($logsnapvol -eq $dbsnapvol))

        $nl
    }

    Write-Debug "dbsnapvol: $dbsnapvol | logsnapvol: $logsnapvol"

    # expose the drives
    # if volumes are the same only one entry is needed
    if ($dbEdbVol -eq $dbLogVol) {
        if ($dbEdbVol.length -gt "2") {
            $dbvolstr = "expose %vss_test_" + ($dbEdbVol).substring(11, 8) + "% $dbsnapvol"
            Out-DHSFile $dbvolstr
        } else {
            $dbvolstr = "expose %vss_test_" + ($dbEdbVol).substring(0, 1) + "% $dbsnapvol"
            Out-DHSFile $dbvolstr
        }
    } else {
        # volumes are different, getting both
        # if mountpoint use first part of string, if not use first letter
        if ($dbEdbVol.length -gt "2") {
            $dbvolstr = "expose %vss_test_" + ($dbEdbVol).substring(11, 8) + "% $dbsnapvol"
            Out-DHSFile $dbvolstr
        } else {
            $dbvolstr = "expose %vss_test_" + ($dbEdbVol).substring(0, 1) + "% $dbsnapvol"
            Out-DHSFile $dbvolstr
        }

        # if mountpoint use first part of string, if not use first letter
        if ($dbLogVol.length -gt "2") {
            $logvolstr = "expose %vss_test_" + ($dbLogVol).substring(11, 8) + "% $logsnapvol"
            Out-DHSFile $logvolstr
        } else {
            $logvolstr = "expose %vss_test_" + ($dbLogVol).substring(0, 1) + "% $logsnapvol"
            Out-DHSFile $logvolstr
        }
    }

    # ending data of file
    Out-DHSFile "end backup"
}

function Invoke-DiskShadow {
    Write-Host " " $nl
    Get-Date
    Write-Host "Starting DiskShadow copy of Exchange database: $selDB" -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    Write-Host "Running the following command:" $nl
    Write-Host "`"C:\Windows\System32\diskshadow.exe /s $path\diskshadow.dsh /l $path\diskshadow.log`"" $nl
    Write-Host " "

    #in case the $path and the script location is different we need to change location into the $path directory to get the results to work as expected.
    try {
        $here = (Get-Location).Path
        Set-Location $path
        diskshadow.exe /s $path\diskshadow.dsh /l $path\diskshadow.log
    } finally {
        Set-Location $here
    }
}

function Invoke-RemoveExposedDrives {

    function Out-removeDHSFile {
        param ([string]$fileline)
        $fileline | Out-File -FilePath "$path\removeSnapshot.dsh" -Encoding ASCII -Append
    }

    " "
    Get-Date
    Write-Host "Diskshadow Snapshots" -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    Write-Host " "
    if ($null -eq $logsnapvol) {
        $exposedDrives = $dbsnapvol
    } else {
        $exposedDrives = $dbsnapvol.ToString() + " and " + $logsnapvol.ToString()
    }
    "If the snapshot was successful, the snapshot should be exposed as drive(s) $exposedDrives."
    "You should be able to see and navigate the snapshot with File Explorer. How would you like to proceed?"
    Write-Host " "
    Write-Host "NOTE: It is recommended to wait a few minutes to allow truncation to possibly occur before moving past this point." -ForegroundColor Cyan
    Write-Host "      This allows time for the logs that are automatically collected to include the window for the truncation to occur." -ForegroundColor Cyan
    Write-Host
    Write-Host "When ready, choose from the options below:" -ForegroundColor Yellow
    " "
    Write-Host "  1. Remove exposed snapshot now"
    Write-Host "  2. Keep snapshot exposed"
    Write-Host " "
    Write-Warning "Selecting option 1 will permanently delete the snapshot created, i.e. your backup will be deleted."
    " "
    $matchCondition = "^[1-2]$"
    Write-Debug "matchCondition: $matchCondition"
    do {
        Write-Host "Selection" -ForegroundColor Yellow -NoNewline
        $removeExpose = Read-Host " "
        if ($removeExpose -notmatch $matchCondition) {
            Write-Host "Error! Please choose a valid option." -ForegroundColor red
        }
    } while ($removeExpose -notmatch $matchCondition)

    $unexposeCommand = "delete shadows exposed $dbsnapvol"
    if ($null -ne $logsnapvol) {
        $unexposeCommand += $nl + "delete shadows exposed $logsnapvol"
    }

    if ($removeExpose -eq "1") {
        New-Item -Path $path\removeSnapshot.dsh -type file -Force
        Out-removeDHSFile $unexposeCommand
        Out-removeDHSFile "exit"
        & 'C:\Windows\System32\diskshadow.exe' /s $path\removeSnapshot.dsh
    } elseif ($removeExpose -eq "2") {
        Write-Host "You can remove the snapshots at a later time using the diskshadow tool from a command prompt."
        Write-Host "Run diskshadow followed by these commands:"
        Write-Host $unexposeCommand
    }
}

function Get-CopyStatus {
    if ((($databases[$dbToBackup]).IsMailboxDatabase) -eq "True") {
        Get-Date
        Write-Host "Status of '$selDB' and its replicas (if any)" -ForegroundColor Green $nl
        Write-Host "--------------------------------------------------------------------------------------------------------------"
        " "
        [array]$copyStatus = (Get-MailboxDatabaseCopyStatus -identity ($databases[$dbToBackup]).name)
        ($copyStatus | Format-List) | Out-File -FilePath "$path\copyStatus.txt"
        for ($i = 0; $i -lt ($copyStatus).length; $i++ ) {
            if (($copyStatus[$i].status -eq "Healthy") -or ($copyStatus[$i].status -eq "Mounted")) {
                Write-Host "$($copyStatus[$i].name) is $($copyStatus[$i].status)"
            } else {
                Write-Host "$($copyStatus[$i].name) is $($copyStatus[$i].status)"
                Write-Host "One of the copies of the selected database is not healthy. Please run backup after ensuring that the database copy is healthy" -ForegroundColor Yellow
                exit
            }
        }
    } Else {
        Write-Host "Not checking database copy status since the selected database is a Public Folder Database..."
    }
    " "
}

function Get-Databases {
    Get-Date
    Write-Host "Getting databases on server: $serverName" -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "

    [array]$script:databases = Get-MailboxDatabase -server $serverName -status
    if ($null -ne (Get-PublicFolderDatabase -Server $serverName)) {
        $script:databases += Get-PublicFolderDatabase -server $serverName -status
    }

    #write-host "Database Name:`t`t Mounted: `t`t Mounted On Server:" -foregroundcolor Yellow $nl
    $script:dbID = 0

    foreach ($script:db in $databases) {
        $script:db | Add-Member NoteProperty Number $dbID
        $dbID++
    }

    $script:databases | Format-Table Number, Name, Mounted, Server -AutoSize | Out-String

    Write-Host " " $nl
}

function Get-DBtoBackup {
    $maxDbIndexRange = $script:databases.length - 1
    $matchCondition = "^([0-9]|[1-9][0-9])$"
    Write-Debug "matchCondition: $matchCondition"
    do {
        Write-Host "Select the number of the database to backup" -ForegroundColor Yellow -NoNewline;
        $script:dbToBackup = Read-Host " "

        if ($script:dbToBackup -notmatch $matchCondition -or [int]$script:dbToBackup -gt $maxDbIndexRange) {
            Write-Host "Error! Please select a valid option!" -ForegroundColor Red
        }
    } while ($script:dbToBackup -notmatch $matchCondition -or [int]$script:dbToBackup -gt $maxDbIndexRange) # notmatch is case-insensitive

    if ((($databases[$dbToBackup]).IsMailboxDatabase) -eq "True") {

        $script:dbGuid = (Get-MailboxDatabase ($databases[$dbToBackup])).guid
        $script:selDB = (Get-MailboxDatabase ($databases[$dbToBackup])).name
        " "
        "The database guid for '$selDB' is: $dbGuid"
        " "
        $script:dbMountedOn = (Get-MailboxDatabase ($databases[$dbToBackup])).server.name
    } else {
        $script:dbGuid = (Get-PublicFolderDatabase ($databases[$dbToBackup])).guid
        $script:selDB = (Get-PublicFolderDatabase ($databases[$dbToBackup])).name
        "The database guid for '$selDB' is: $dbGuid"
        " "
        $script:dbMountedOn = (Get-PublicFolderDatabase ($databases[$dbToBackup])).server.name
    }
    Write-Host "The database is mounted on server: $dbMountedOn $nl"

    if ($dbMountedOn -eq "$serverName") {
        $script:dbStatus = "active"
    } else {
        $script:dbStatus = "passive"
    }
}

function Get-ExchangeVersion {
    Get-Date
    Write-Host "Verifying Exchange version..." -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    $script:exchVer = (Get-ExchangeServer $serverName).AdminDisplayVersion
    $exchVerMajor = $exchVer.major
    $exchVerMinor = $exchVer.minor

    switch ($exchVerMajor) {
        "14" {
            $script:exchVer = "2010"
        }
        "15" {
            switch ($exchVerMinor) {
                "0" {
                    $script:exchVer = "2013"
                }
                "1" {
                    $script:exchVer = "2016"
                }
                "2" {
                    $script:exchVer = "2019"
                }
            }
        }

        default {
            Write-Host "This script is only for Exchange 2013, 2016, and 2019 servers." -ForegroundColor red $nl
            exit
        }
    }

    Write-Host "$serverName is an Exchange $exchVer server. $nl"

    if ($exchVer -eq "2010") {
        Write-Host "This script no longer supports Exchange 2010."
        exit
    }
}


function Get-WindowsEventLogs {

    Function Get-WindowEventsPerServer {
        param(
            [string]$CompunterName
        )
        "Getting application log events..."
        Get-WinEvent -FilterHashtable @{LogName = "Application"; StartTime = $startTime } -ComputerName $CompunterName | Export-Clixml $path\events-$CompunterName-App.xml
        "Getting system log events..."
        Get-WinEvent -FilterHashtable @{LogName = "System"; StartTime = $startTime } -ComputerName $CompunterName | Export-Clixml $path\events-$CompunterName-System.xml
        "Getting TruncationDebug log events..."
        Get-WinEvent -FilterHashtable @{LogName = "Microsoft-Exchange-HighAvailability/TruncationDebug"; StartTime = $startTime } -ComputerName $CompunterName -ErrorAction SilentlyContinue | Export-Clixml $path\events-$CompunterName-TruncationDebug.xml
    }
    " "
    Get-Date
    Write-Host "Getting events from the application and system logs since the script's start time of ($startInfo)" -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    "Getting application log events..."
    Get-EventLog -LogName Application -After $startInfo | Export-Clixml $path\events-App.xml
    "Getting system log events..."
    Get-EventLog -LogName System -After $startInfo | Export-Clixml $path\events-System.xml
    "Getting events complete!"
}

function Get-VSSWritersAfter {
    " "
    Get-Date
    Write-Host "Checking VSS Writer Status: (after backup)" -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    " "
    $writers = (vssadmin list writers)
    $writers > $path\vssWritersAfter.txt

    foreach ($line in $writers) {
        if ($line -like "Writer name:*") {
            "$line"
        } elseif ($line -like "   State:*") {
            "$line" + $nl
        }
    }
}

function Get-VSSWritersBefore {
    " "
    Get-Date
    Write-Host "Checking VSS Writer Status: (All Writers must be in a Stable state before running this script)" -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    $writers = (vssadmin list writers)
    $writers > $path\vssWritersBefore.txt
    $exchangeWriter = $false

    foreach ($line in $writers) {

        if ($line -like "Writer name:*") {
            "$line"

            if ($line.Contains("Microsoft Exchange Writer")) {
                $exchangeWriter = $true
            }
        } elseif ($line -like "   State:*") {

            if ($line -ne "   State: [1] Stable") {
                $nl
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!   WARNING   !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor red
                $nl
                Write-Host "One or more writers are NOT in a 'Stable' state, STOPPING SCRIPT." -ForegroundColor red
                $nl
                Write-Host "Review the vssWritersBefore.txt file in '$path' for more information." -ForegroundColor Red
                Write-Host "You can also use an Exchange Management Shell or a Command Prompt to run: 'vssadmin list writers'" -ForegroundColor red
                $nl
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor red
                $nl
                exit
            } else {
                "$line" + $nl
            }
        }
    }
    " " + $nl

    if (!$exchangeWriter) {

        #Check for possible COM security issue.
        $oleKey = "HKLM:\SOFTWARE\Microsoft\Ole"
        $dcomKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DCOM"

        $possibleDcomPermissionIssue = ((Test-Path $oleKey) -and
            ($null -ne (Get-ItemProperty $oleKey).DefaultAccessPermission)) -or
        ((Test-Path $dcomKey) -and
            ($null -ne (Get-ItemProperty $dcomKey).MachineAccessRestriction))

        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!   WARNING   !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor red
        Write-Host "Microsoft Exchange Writer not present on server. Unable to preform proper backups on the server."  -ForegroundColor Red
        Write-Host

        if ($possibleDcomPermissionIssue) {
            Write-Host " - Recommend to verify local Administrators group applied to COM+ Security settings: https://aka.ms/VSSTester-COMSecurity" -ForegroundColor Cyan
        }

        Write-Host " - Recommend to restart MSExchangeRepl service to see if the writer comes back. If it doesn't, review the application logs for any events to determine why." -ForegroundColor Cyan
        Write-Host " --- Look for Event ID 2003 in the application logs to verify that all internal components come online. If you see this event, try to use PSExec.exe to start a cmd.exe as the SYSTEM account and run 'vssadmin list writers'" -ForegroundColor Cyan
        Write-Host " --- If you find the Microsoft Exchange Writer, then we have a permissions issue on the computer that is preventing normal user accounts from finding all the writers." -ForegroundColor Cyan
        Write-Host " - If still not able to determine why, need to have a Microsoft Engineer review ExTrace with Cluster.Replay tags of the MSExchangeRepl service starting up." -ForegroundColor Cyan
        Write-Host
        Write-Host "Stopping Script"
        exit
    }
}

function Invoke-CreateExTRATracingConfig {

    function Out-ExTRAConfigFile {
        param ([string]$fileline)
        $fileline | Out-File -FilePath "C:\EnabledTraces.Config" -Encoding ASCII -Append
    }

    " "
    Get-Date
    Write-Host "Enabling ExTRA Tracing..." -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    New-Item -Path "C:\EnabledTraces.Config" -type file -Force

    Out-ExTRAConfigFile "TraceLevels:Debug,Warning,Error,Fatal,Info,Performance,Function,Pfd"
    Out-ExTRAConfigFile "ManagedStore.PhysicalAccess:JetBackup,JetRestore,JetEventlog,SnapshotOperation"
    Out-ExTRAConfigFile "Cluster.Replay:LogTruncater,ReplayApi,ReplicaInstance,ReplicaVssWriterInterop"
    Out-ExTRAConfigFile "ManagedStore.HA:BlockModeSender,Eseback"
    Out-ExTRAConfigFile "FilteredTracing:No"
    Out-ExTRAConfigFile "InMemoryTracing:No"
    " "
    Write-Debug "ExTRA trace config file created successfully"
}

function Invoke-DisableDiagnosticsLogging {

    Write-Host " "  $nl
    Get-Date
    Write-Host "Disabling Diagnostics Logging..." -ForegroundColor green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    Set-EventLogLevel 'MSExchange Repl\Service' -level lowest
    $disgetReplSvc = Get-EventLogLevel 'MSExchange Repl\Service'
    Write-Host "$($disgetReplSvc.Identity) - $($disgetReplSvc.EventLevel) $nl"

    Set-EventLogLevel 'MSExchange Repl\Exchange VSS Writer' -level lowest
    $disgetReplVSSWriter = Get-EventLogLevel 'MSExchange Repl\Exchange VSS Writer'
    Write-Host "$($disgetReplVSSWriter.Identity) - $($disgetReplVSSWriter.EventLevel) $nl"
}

function Invoke-DisableExTRATracing {
    " "
    Get-Date
    Write-Host "Disabling ExTRA Tracing..." -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    if ($dbMountedOn -eq "$serverName") {
        #stop active copy
        Write-Host " "
        "Stopping Exchange Trace data collector on $serverName..."
        logman stop vssTester -s $serverName
        "Deleting Exchange Trace data collector on $serverName..."
        logman delete vssTester -s $serverName
        " "
    } else {
        #stop passive copy
        "Stopping Exchange Trace data collector on $serverName..."
        logman stop vssTester-Passive -s $serverName
        "Deleting Exchange Trace data collector on $serverName..."
        logman delete vssTester-Passive -s $serverName
        #stop active copy
        "Stopping Exchange Trace data collector on $dbMountedOn..."
        logman stop vssTester-Active -s $dbMountedOn
        "Deleting Exchange Trace data collector on $dbMountedOn..."
        logman delete vssTester-Active -s $dbMountedOn
        " "
        "Moving ETL file from $dbMountedOn to $serverName..."
        " "
        $etlPath = $path -replace ":\\", "$\"
        Move-Item "\\$dbMountedOn\$etlPath\vsstester-active_000001.etl" "\\$servername\$etlPath\vsstester-active_000001.etl" -Force
    }
}

function Invoke-DisableVSSTracing {
    " "
    Get-Date
    Write-Host "Disabling VSS Tracing..." -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    logman stop vss -ets
    " "
}

function Invoke-EnableDiagnosticsLogging {
    " "
    Get-Date
    Write-Host "Enabling Diagnostics Logging..." -ForegroundColor green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    Set-EventLogLevel 'MSExchange Repl\Service' -level expert
    $getReplSvc = Get-EventLogLevel 'MSExchange Repl\Service'
    Write-Host "$($getReplSvc.Identity) - $($getReplSvc.EventLevel) $nl"

    Set-EventLogLevel 'MSExchange Repl\Exchange VSS Writer' -level expert
    $getReplVSSWriter = Get-EventLogLevel 'MSExchange Repl\Exchange VSS Writer'
    Write-Host "$($getReplVSSWriter.Identity)  - $($getReplVSSWriter.EventLevel)  $nl"
}

function Invoke-EnableExTRATracing {

    Function Invoke-ExtraTracingCreate {
        param(
            [string]$ComputerName,
            [string]$LogmanName
        )
        [array]$results = logman create trace $LogmanName -p '{79bb49e6-2a2c-46e4-9167-fa122525d540}' -o $path\$LogmanName.etl -ow -s $ComputerName -mode globalsequence
        $results

        if ($results[-1] -eq "Data Collector already exists.") {
            Write-Host "Exchange Trace data Collector set already created. Removing it and trying again"
            [array]$results = logman delete $LogmanName -s $ComputerName
            $results

            [array]$results = logman create trace $LogmanName -p '{79bb49e6-2a2c-46e4-9167-fa122525d540}' -o $path\$LogmanName.etl -ow -s $ComputerName -mode globalsequence
            $results
        }

        if ($results[-1] -ne "The command completed successfully.") {
            Write-Host "Failed to create the extra trace. Stopping the VSSTester Script" -ForegroundColor Red
            exit
        }
    }

    #active server, only get tracing from active node
    if ($dbMountedOn -eq $serverName) {
        " "
        "Creating Exchange Trace data collector set..."
        Invoke-ExtraTracingCreate -ComputerName $serverName -LogmanName "VSSTester"
        "Starting Exchange Trace data collector..."
        [array]$results = logman start VSSTester
        $results

        if ($results[-1] -ne "The command completed successfully.") {
            Write-Host "Failed to start the extra trace. Stopping the VSSTester Script" -ForegroundColor Red
            exit
        }
        " "
    } else {
        #passive server, get tracing from both active and passive nodes
        " "
        "Copying the ExTRA config file 'EnabledTraces.config' file to $dbMountedOn..."
        #copy enabledtraces.config from current passive copy to active copy server
        Copy-Item "c:\EnabledTraces.Config" "\\$dbMountedOn\c$\enabledtraces.config" -Force

        #create trace on passive copy
        "Creating Exchange Trace data collector set on $serverName..."
        Invoke-ExtraTracingCreate -ComputerName $serverName -LogmanName "VSSTester-Passive"
        #create trace on active copy
        "Creating Exchange Trace data collector set on $dbMountedOn..."
        Invoke-ExtraTracingCreate -ComputerName $dbMountedOn -LogmanName "VSSTester-Active"
        #start trace on passive copy
        "Starting Exchange Trace data collector on $serverName..."
        [array]$results = logman start VSSTester-Passive -s $serverName
        $results

        if ($results[-1] -ne "The command completed successfully.") {
            Write-Host "Failed to start the extra trace. Stopping the VSSTester Script" -ForegroundColor Red
            exit
        }
        #start trace on active copy
        "Starting Exchange Trace data collector on $dbMountedOn..."
        [array]$results = logman start VSSTester-Active -s $dbMountedOn
        $results

        if ($results[-1] -ne "The command completed successfully.") {
            Write-Host "Failed to start the extra trace. Stopping the VSSTester Script" -ForegroundColor Red
            exit
        }
        " "
    }

    Write-Debug "ExTRA trace started successfully"
}

function Invoke-EnableVSSTracing {
    " "
    Get-Date
    Write-Host "Enabling VSS Tracing..." -ForegroundColor Green $nl
    Write-Host "--------------------------------------------------------------------------------------------------------------"
    " "
    logman start vss -o $path\vss.etl -ets -p "{9138500e-3648-4edb-aa4c-859e9f7b7c38}" 0xfff 255
}

Function Main {

    # if a transcript is running, we need to stop it as this script will start its own
    try {
        Stop-Transcript | Out-Null
    } catch [System.InvalidOperationException] { }

    Write-Host "****************************************************************************************"
    Write-Host "****************************************************************************************"
    Write-Host "**                                                                                    **" -BackgroundColor DarkMagenta
    Write-Host "**                 VSSTESTER SCRIPT (for Exchange 2013, 2016, 2019)                   **" -ForegroundColor Cyan -BackgroundColor DarkMagenta
    Write-Host "**                                                                                    **" -BackgroundColor DarkMagenta
    Write-Host "****************************************************************************************"
    Write-Host "****************************************************************************************"

    $Script:LocalExchangeShell = Confirm-ExchangeShell -Identity $env:COMPUTERNAME

    if (!$Script:LocalExchangeShell.ShellLoaded) {
        Write-Host "Failed to load Exchange Shell. Stopping the script."
        exit
    }

    if ($Script:LocalExchangeShell.RemoteShell -or
        $Script:LocalExchangeShell.ToolsOnly) {
        Write-Host "Can't run this script from a non Exchange Server."
        exit
    }

    #newLine shortcut
    $script:nl = "`r`n"
    $nl

    $script:serverName = $env:COMPUTERNAME

    #start time
    $Script:startInfo = Get-Date
    Get-Date

    if ($DebugPreference -ne 'SilentlyContinue') {
        $nl
        Write-Host 'This script is running in DEBUG mode since $DebugPreference is not set to SilentlyContinue.' -ForegroundColor Red
    }

    $nl
    Write-Host "Please select the operation you would like to perform from the following options:" -ForegroundColor Green
    $nl
    Write-Host "  1. " -ForegroundColor Yellow -NoNewline; Write-Host "Test backup using built-in Diskshadow"
    Write-Host "  2. " -ForegroundColor Yellow -NoNewline; Write-Host "Enable logging to troubleshoot backup issues"
    $nl

    $matchCondition = "^[1|2]$"
    Write-Debug "matchCondition: $matchCondition"
    Do {
        Write-Host "Selection: " -ForegroundColor Yellow -NoNewline;
        $Selection = Read-Host
        if ($Selection -notmatch $matchCondition) {
            Write-Host "Error! Please select a valid option!" -ForegroundColor Red
        }
    }
    while ($Selection -notmatch $matchCondition)


    try {

        $nl
        Write-Host "Please specify a directory other than root of a volume to save the configuration and output files." -ForegroundColor Green

        $pathExists = $false

        # get path, ensuring it exists
        do {
            Write-Host "Directory path (e.g. C:\temp): " -ForegroundColor Yellow -NoNewline
            $script:path = Read-Host
            Write-Debug "path: $path"
            try {
                $pathExists = Test-Path -Path "$path"
            } catch { }
            Write-Debug "pathExists: $pathExists"
            if ($pathExists -ne $true) {
                Write-Host "Error! The path does not exist. Please enter a valid path." -ForegroundColor red
            }
        } while ($pathExists -ne $true)

        $nl
        Get-Date
        Write-Host "Starting transcript..." -ForegroundColor Green $nl
        Write-Host "--------------------------------------------------------------------------------------------------------------"

        Start-Transcript -Path "$($script:path)\vssTranscript.log"
        $nl

        if ($Selection -eq 1) {
            Get-ExchangeVersion
            Get-VSSWritersBefore
            Get-Databases
            Get-DBtoBackup
            Get-CopyStatus
            Invoke-CreateDiskShadowFile #---
            Invoke-EnableDiagnosticsLogging
            Invoke-EnableVSSTracing
            Invoke-CreateExTRATracingConfig
            Invoke-EnableExTRATracing
            Invoke-DiskShadow #---
            Get-VSSWritersAfter
            Invoke-RemoveExposedDrives #---
            Invoke-DisableExTRATracing
            Invoke-DisableDiagnosticsLogging
            Invoke-DisableVSSTracing
            Get-WindowsEventLogs
        } elseif ($Selection -eq 2) {
            Get-ExchangeVersion
            Get-VSSWritersBefore
            Get-Databases
            Get-DBtoBackup
            Get-CopyStatus
            Invoke-EnableDiagnosticsLogging
            Invoke-EnableVSSTracing
            Invoke-CreateExTRATracingConfig
            Invoke-EnableExTRATracing

            #Here is where we wait for the end user to perform the backup using the backup software and then come back to the script to press "Enter", thereby stopping data collection
            Get-Date
            Write-Host "Data Collection" -ForegroundColor green $nl
            Write-Host "--------------------------------------------------------------------------------------------------------------"
            " "
            Write-Host "Data collection is now enabled." -ForegroundColor Yellow
            Write-Host "Please start your backup using the third party software so the script can record the diagnostic data." -ForegroundColor Yellow
            Write-Host "When the backup is COMPLETE, use the <Enter> key to terminate data collection..." -ForegroundColor Yellow -NoNewline
            Read-Host

            Invoke-DisableExTRATracing
            Invoke-DisableDiagnosticsLogging
            Invoke-DisableVSSTracing
            Get-VSSWritersAfter
            Get-WindowsEventLogs
        }
    } finally {
        # always stop our transcript at end of script's execution
        # we catch a failure here if we try to stop a transcript that's not running
        try {
            " " + $nl
            Get-Date
            Write-Host "Stopping transcript log..." -ForegroundColor Green $nl
            Write-Host "--------------------------------------------------------------------------------------------------------------"
            " "
            Stop-Transcript
            " " + $nl
            do {
                Write-Host
                $continue = Read-Host "Please use the <Enter> key to exit..."
            }
            While ($null -notmatch $continue)
            exit
        } catch { }
    }
}

try {
    Clear-Host
    Main
} catch { } finally { }

# SIG # Begin signature block
# MIIjngYJKoZIhvcNAQcCoIIjjzCCI4sCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAUM3zRxRXTQXwR
# A924zkcu7frFIjdqvrrvaZxb47O3CKCCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVczCCFW8CAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBxjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgxGMLVWnI
# Q3uSPIuBA9AlbMEusQKjKhYbR6HsuY5H6HEwWgYKKwYBBAGCNwIBDDFMMEqgGoAY
# AEMAUwBTACAARQB4AGMAaABhAG4AZwBloSyAKmh0dHBzOi8vZ2l0aHViLmNvbS9t
# aWNyb3NvZnQvQ1NTLUV4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQBIVna7jkvp
# g0UwAnvPJJB3tqyiIG+e7usykyQN3rPu9ZEbSL3N7ou2wHQB8PdIFGhRrZD+27sv
# ieyWLQHQ1q/JgWiOV0VfkCLifenk2Z7WBte212+FhylOtS5Tl29Nxh9d91jXltxm
# FyiukT8XS79eAmg+NUKgBveQpyHeW4TgjriRtCA5NWbVxv/h761Lw3wTYWlJRrET
# g8v0n+qwr4wRjPTA4aQ+qWhUnQ9zFT9ons/bNl/eCqAK3AJKtV2dzzmXbWOqkGZV
# SydTMKKxXx4l99Xk+wNKYdQx47aj00yQ2mEOOjc5Ch9wVM9jyvdHwJlEMSoBE8c0
# 9RPp8dTpVNGaoYIS5TCCEuEGCisGAQQBgjcDAwExghLRMIISzQYJKoZIhvcNAQcC
# oIISvjCCEroCAQMxDzANBglghkgBZQMEAgEFADCCAVEGCyqGSIb3DQEJEAEEoIIB
# QASCATwwggE4AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIFIwSoaG
# fH89tE0cbjKq9gaslVtLpgurdsS+llfNWc7jAgZhHpuib2sYEzIwMjEwOTAzMjIw
# MjA5LjAwN1owBIACAfSggdCkgc0wgcoxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjIyNjQtRTMzRS03ODBDMSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIOPDCCBPEwggPZoAMC
# AQICEzMAAAFKpPcxxP8iokkAAAAAAUowDQYJKoZIhvcNAQELBQAwfDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjAxMTEyMTgyNTU4WhcNMjIwMjExMTgy
# NTU4WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMG
# A1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEmMCQGA1UECxMdVGhh
# bGVzIFRTUyBFU046MjI2NC1FMzNFLTc4MEMxJTAjBgNVBAMTHE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDeyihmZKJYLL1RGjSjE1WWYBJfKIbC4B0eIFBVi2b1sy23oA6ESLaXxXfvZmlt
# oTxZYE/sL+5cX+jgeBxWGYB3yKXGYRlOv3m7Mpl2AJgCsyqYe9acSVORdtvGE0ky
# 3KEgCFDQWVXUxCGSCxD0+YCO+2LLu2CjLn0pomT86mJZBF9v3R4TnKKPdM4CCUUx
# tbtpBe8Omuw+dMhyhOOnhhMKsIxMREQgjbRQQ0K032CA/yHI9MyopGI4iUWmjzY5
# 7wWkSf3hZBs/IA9l8mF45bDYwxj2hj0E7f0Zt568XMlxsgiCIVnQTFzEy5ewTAyi
# niwUNHeqRX0tS0SaPqWiigYlAgMBAAGjggEbMIIBFzAdBgNVHQ4EFgQUcYxhGDH6
# wIY1ipP/fX64JiqpP+EwHwYDVR0jBBgwFoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUw
# VgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9j
# cmwvcHJvZHVjdHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3JsMFoGCCsGAQUF
# BwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIw
# ADATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQsFAAOCAQEAUQuu3UY4
# BRUvZL+9lX3vIEPh4NxaV9k2MjquJ67T6vQ9+lHcna9om2cuZ+y6YV71ttGw07oF
# B4sLsn1p5snNqBHr6PkqzQs8V3I+fVr/ZUKQYLS+jjOesfr9c2zc6f5qDMJN1L8r
# BOWn+a5LXxbT8emqanI1NSA7dPYV/NGQM6j35Tz8guQo9yfA0IpUM9v080mb3G4A
# jPb7sC7vafW2YSXpjT/vty6x5HcnHx2X947+0AQIoBL8lW9pq55aJhSCgsiVtXDq
# wYyKsp7ULeTyvMysV/8mZcokW6/HNA0MPLWKV3sqK4KFXrfbABfrd4P3GM1aIFuK
# sIbsmZhJk5U0ijCCBnEwggRZoAMCAQICCmEJgSoAAAAAAAIwDQYJKoZIhvcNAQEL
# BQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNV
# BAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4X
# DTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIxNDY1NVowfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28
# dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF++18aEssX8XD5WHCdrc+Zitb8BVTJwQx
# H0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRDDNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVH
# gc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSxz5NMksHEpl3RYRNuKMYa+YaAu99h/EbB
# Jx0kZxJyGiGKr0tkiVBisV39dx898Fd1rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL
# /W7lmsqxqPJ6Kgox8NpOBpG2iAg16HgcsOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8
# wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNV
# HQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqFbVUwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwgaAGA1UdIAEB/wSBlTCBkjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsG
# AQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vUEtJL2RvY3MvQ1BTL2Rl
# ZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAFAAbwBsAGkA
# YwB5AF8AUwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH
# 5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUxvs8F4qn++ldtGTCzwsVmyWrf9efweL3H
# qJ4l4/m87WtUVwgrUYJEEvu5U4zM9GASinbMQEBBm9xcF/9c+V4XNZgkVkt070IQ
# yK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1L3mBZdmptWvkx872ynoAb0swRCQiPM/t
# A6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWOM7tiX5rbV0Dp8c6ZZpCM/2pif93FSguR
# JuI57BlKcWOdeyFtw5yjojz6f32WapB4pm3S4Zz5Hfw42JT0xqUKloakvZ4argRC
# g7i1gJsiOCC1JeVk7Pf0v35jWSUPei45V3aicaoGig+JFrphpxHLmtgOR5qAxdDN
# p9DvfYPw4TtxCd9ddJgiCGHasFAeb73x4QDf5zEHpJM692VHeOj4qEir995yfmFr
# b3epgcunCaw5u+zGy9iCtHLNHfS4hQEegPsbiSpUObJb2sgNVZl6h3M7COaYLeqN
# 4DMuEin1wC9UJyH3yKxO2ii4sanblrKnQqLJzxlBTeCG+SqaoxFmMNO7dDJL32N7
# 9ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zP
# WAUu7w2gUDXa7wknHNWzfjUeCLraNtvTX4/edIhJEqGCAs4wggI3AgEBMIH4oYHQ
# pIHNMIHKMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYD
# VQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMSYwJAYDVQQLEx1UaGFs
# ZXMgVFNTIEVTTjoyMjY0LUUzM0UtNzgwQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUAvATuhoUgysEzdykE1bRB
# 4oh6a5iggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkq
# hkiG9w0BAQUFAAIFAOTc4BMwIhgPMjAyMTA5MDQwMTU1MzFaGA8yMDIxMDkwNTAx
# NTUzMVowdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA5NzgEwIBADAKAgEAAgIF2gIB
# /zAHAgEAAgIROTAKAgUA5N4xkwIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEE
# AYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GB
# AI2KMp5FkaW17LODLUsGPK7iplUaXdveKoD63p9ct5TnMa02BHXxAcniWJGFChjf
# sOUvkAnPGsFW7vWpnqRBMrvWKiprzTRZwE9q6KrQNBcuUz6hw3wJqJEbfyBGAmsv
# IQWDYqP+Y9Kupino7Nv7L5scrMVvaylUvGBfew+xLavxMYIDDTCCAwkCAQEwgZMw
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAFKpPcxxP8iokkAAAAA
# AUowDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRAB
# BDAvBgkqhkiG9w0BCQQxIgQgs7wELuJWmKBDwYcaJ5ygxuk53nlthBJ5jZsJlYlU
# MgYwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCBsHZLXrbnbV/5J+2KvwFWI
# gVmQavp+BBVUPM1A9yJRAzCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAABSqT3McT/IqJJAAAAAAFKMCIEIByhRiX1VAytHWXmtYDFKCDg
# FJ/sjaN/5ONVQHRUz+fYMA0GCSqGSIb3DQEBCwUABIIBAGitObO3rFse3TLz7G5b
# 7dzh7ohogfWiwzMN8Q9vTYLrIly/oxPU3rm89XaGwwrJbA0SEEmt1BUbgVroAnUM
# no7Y1AqxEPLOfGE3m5daqqgW+LB3WHAI7WPLnls5HjCprmLLhD1hK+6n06DH62Ah
# 87DRDH+5wG6SUbRFgcRmpox57iJ/1fWg7v+BOC9uzFDtiyHEowAcereqntCxl/La
# RoSrRE1X0+zGDNR+8tMlnfwIvjTVCcXWJMYyMaAUX8hYBwodGACUW7MziB1LpHe0
# 3IlO02gNfV4ni2xqf3n125dTs3eIz1dp0txSAFHM/qUt+EwwZe3jp7EOyfT2AfbQ
# IdI=
# SIG # End signature block
