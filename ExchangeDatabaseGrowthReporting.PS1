#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################

#Authors : Jesse Newgard , Charles Lewis

#################################################################################
# This script will allow you to troubleshoot database growth issue by collecting the necessary data to determine where the growth is coming from.
# The script can only be used with Exchange 2010 and Exchange 2013
# The script requires powershell 2.0 or later
#
#################################################################################

#Script Examples:



Param(
	[ValidateRange(0,7)][int]$Mode = 0
#	,[bool]$Test = 0
	)


##################################################Misc. Functions##################################################

#####Convert-Size#####
#[String]
#Takes a size in Bytes and outputs that size in the smallest appropriate magnitude to and to two decimal places.
######################

Function Convert-Size ($size)
{
	if ($size -lt 1024) {write "$($size) B"}
	elseif ($size -ge 1024 -and $size -lt 1MB) {write "$("{0:N2}" -f ($size / 1KB)) KB"}
	elseif ($size -ge 1MB -and $size -lt 1GB) {write "$("{0:N2}" -f ($size / 1MB)) MB"}
	elseif ($size -ge 1GB -and $size -lt 1024GB) {write "$("{0:N2}" -f ($size / 1GB)) GB"}
	else {write "$("{0:N2}" -f ($size / 1024GB)) TB"}
}

#####Get-TimeStamp####
#[Void]
#Gathers the current M/D/Y in addition to current System Ticks to allow for multiple reports on the same date and outputs a combined string "-M-D-Y-Ticks"
######################

Function Get-TimeStamp ()
{
    $date = Get-Date
    
    $timeStamp = '-' + $Date.Month.ToString() + '-' + $Date.Day.ToString() + '-' + $Date.Year.ToString() + '-' + $Date.Ticks.ToString()  
	
	$timeStamp
}
##################################################/Misc Functions####################################################


##################################################Collection Functions###############################################
####Get-mbxStatistics#
#[Mailbox]
#Collects mailbox statistics across the specified database, adding mutable note properties for future use in differencing
######################

Function Get-mbxStatistics ($dbName)
{
	If ((Get-ExchangeServer(Get-MailboxDatabase $dbName).Server).AdminDisplayVersion -NotLike "Version 15*")
	{
		Clean-MailboxDatabase $dbName
	}

	[System.Array]$mbxStat = Get-MailboxStatistics -Database $dbName
	$mbxStat | %{Add-Member -InputObject $_ -MemberType NoteProperty -Name DeletedItemSize -Value $_.TotalDeletedItemSize.Value.ToBytes(); Add-Member -InputObject $_ -MemberType NoteProperty -Name ItemSize -Value $_.TotalItemSize.Value.ToBytes(); Add-Member -InputObject $_ -MemberType NoteProperty -Name Items -Value $_.ItemCount; Add-Member -InputObject $_ -MemberType NoteProperty -Name AssociatedItems -Value $_.AssociatedItemCount; Add-Member -InputObject $_ -MemberType NoteProperty -Name DeletedItems -Value $_.DeletedItemCount}
	
	$mbxStat = $mbxStat | Select ItemSize,DeletedItemSize,Items,DeletedItems,AssociatedItems,Database,DisconnectDate,DisconnectReason,Identity,DisplayName,ObjectClass,LegacyDN
	
	$mbxStat
}

####Get-dbStatistics##
#[Database]
#Collects database statistics for the specified database, adding mutable note properties for later differencing.
######################

Function Get-dbStatistics ($dbName)
{
	$dataBase = Get-MailboxDatabase $dbName -Status
	$dataBase | Add-Member –MemberType NoteProperty –Name PhysicalSize –Value $dataBase.DatabaseSize.ToBytes()
	$dataBase | Add-Member –MemberType NoteProperty –Name WhiteSpace –Value $dataBase.AvailableNewMailboxSpace.ToBytes()
	
	#
	
	$mbxStat = Get-mbxStatistics $dbName
	$mbxItemSize = 0
	$mbxDeletedItemSize = 0
	$mbxStat | %{$mbxItemSize += $_.ItemSize; $mbxDeletedItemSize += $_.DeletedItemSize}
	
	$connectedMailboxSize = 0
	$connectedMbx = $mbxStat | ?{$_.DisconnectReason -eq $null}
	$connectedMbx | %{$connectedMailboxSize += ($_.ItemSize + $_.DeletedItemSize)}
	$disconnectedMailboxSize = 0
	$disconnectedMbx = $mbxStat | ?{$_.DisconnectReason -ne $null}
	$disconnectedMbx | %{$disconnectedMailboxSize += ($_.ItemSize + $_.DeletedItemSize)}
	
	#
	
	$dataBase | Add-Member –MemberType NoteProperty –Name ItemSize –Value $mbxItemSize
	$dataBase | Add-Member –MemberType NoteProperty –Name DeletedItemSize –Value $mbxDeletedItemSize
	$dataBase | Add-Member –MemberType NoteProperty –Name ConnectedMailboxSize –Value $connectedMailboxSize
	$dataBase | Add-Member –MemberType NoteProperty –Name DisconnectedMailboxSize –Value $disconnectedMailboxSize
	
	$dataBase | Add-Member –MemberType NoteProperty –Name ActualOverhead –Value ($database.PhysicalSize - ($dataBase.WhiteSpace + $dataBase.ItemSize + $dataBase.DeletedItemSize))
	$dataBase | Add-Member –MemberType NoteProperty –Name ExpectedOverhead –Value ([MATH]::Floor($dataBase.PhysicalSize * 0.2))
	
	$dataBase = $dataBase | Select ItemSize,DeletedItemSize,ActualOverHead,ExpectedOverhead,PhysicalSize,WhiteSpace,Guid,Identity,Name,ConnectedMailboxSize,DisconnectedMailboxSize
	
	$dataBase
}

#Get-mbxFolderStatistics#
#[Folder]
#Collects mailbox folder statistics for all mailboxes on the specified database, adding mutable properties for later differencing.
#########################

Function Get-mbxFolderStatistics ($dbName)
{
	#Since errors/warnings are expected; just hiding the output on black for now
	#$defaultErrorForeGroundColor = $Host.PrivateData.ErrorForeGroundColor
	#$Host.PrivateData.ErrorForeGroundColor = $Host.PrivateData.ErrorBackGroundColor
	#$Host.PrivateData.WarningForeGroundColor = "Black"
	
	
	$mbxStat = Get-mbxStatistics $dbName
	#$mbxStat = $mbxStat | ?{$_.ObjectClass -notlike "*Disabled*" -and $_.ObjectClass -notlike "*System*"}
	[System.Array]$mbxFolderStat = $mbxStat | %{Get-MailboxFolderStatistics $_.Identity.ToString() -ErrorAction SilentlyContinue -WarningAction SilentlyContinue}
	$mbxFolderStat | %{Add-Member -InputObject $_ -MemberType NoteProperty -Name Size -Value $_.FolderSize.ToBytes(); Add-Member -InputObject $_ -MemberType NoteProperty -Name ItemCount -Value $_.ItemsInFolder; Add-Member -InputObject $_ -MemberType NoteProperty -Name DeletedItemCount -Value $_.DeletedItemsInFolder}

	$mbxFolderStat = $mbxFolderStat | Select Identity,Size,ItemCount,DeletedItemCount,FolderID,FolderPath
	
	$mbxFolderStat
	
	#Restoring
	#$Host.PrivateData.ErrorForeGroundColor = $defaultErrorForeGroundColor
	#$Host.PrivateData.WarningForeGroundColor = "Yellow"
}
##################################################/Collection Functions###############################################

#################################################Differencing Functions###############################################
####Compare-Database##
#[Database]
#Subtracts size and item count attributes of the input database from the difference database, returning a database type object with the modified attributes.
######################

Function Compare-Database ($inputDatabase, $differenceDatabase)
{
	$resultantDatabase = $differenceDatabase
	
	$resultantDatabase.PhysicalSize = $resultantDatabase.PhysicalSize - $inputDatabase.PhysicalSize
	$resultantDatabase.ItemSize = $resultantDatabase.ItemSize - $inputDatabase.ItemSize
	$resultantDatabase.DeletedItemSize = $resultantDatabase.DeletedItemSize - $inputDatabase.DeletedItemSize
	$resultantDatabase.WhiteSpace = $resultantDatabase.WhiteSpace - $inputDatabase.WhiteSpace
	
	$resultantDatabase.ConnectedMailboxSize = $resultantDatabase.ConnectedMailboxSize - $inputDatabase.ConnectedMailboxSize
	$resultantDatabase.DisconnectedMailboxSize = $resultantDatabase.DisconnectedMailboxSize - $inputDatabase.DisconnectedMailboxSize	
	
	$resultantDatabase.ActualOverhead = $resultantDatabase.ActualOverhead - $inputDatabase.ActualOverhead
	$resultantDatabase.ExpectedOverhead = $resultantDatabase.ExpectedOverhead - $inputDatabase.ExpectedOverhead
	
	$resultantDatabase
}

####Compare-Mailbox###
#[Mailbox]
#Subtracts size and item count attributes of the input mailbox from the difference mailbox, returning a mailbox type object with the modified attributes.
######################

Function Compare-Mailbox ($inputMailbox, $differenceMailbox)
{
	$resultantMailbox = $differenceMailbox
	
	$resultantMailbox.ItemSize = $resultantMailbox.ItemSize - $inputMailbox.ItemSize
	$resultantMailbox.DeletedItemSize = $resultantMailbox.DeletedItemSize - $inputMailbox.DeletedItemSize
	$resultantMailbox.AssociatedItems = $resultantMailbox.AssociatedItems - $inputMailbox.AssociatedItems
	$resultantMailbox.Items = $resultantMailbox.Items - $inputMailbox.Items
	$resultantMailbox.DeletedItems = $resultantMailbox.DeletedItems - $inputMailbox.DeletedItems
	
	$resultantMailbox
}

####Compare-Folder###
#[Folder]
#Subtracts size and item count attributes of the input folder from the difference folder, returning a folder type object with the modified attributes.
######################

Function Compare-Folder ($inputFolder, $differenceFolder)
{
	$resultantFolder = $differenceFolder
	
	$resultantFolder.Size = $resultantFolder.Size - $inputFolder.Size
	$resultantFolder.ItemCount = $resultantFolder.ItemCount - $inputFolder.ItemCount
	$resultantFolder.DeletedItemCount = $resultantFolder.DeletedItemCount - $inputFolder.DeletedItemCount
	
	$resultantFolder
}


####Compare-Report##
#[Report]
#Subtracts size and item count attributes of the input report from the difference report, returning a report type object with the modified attributes.
######################

Function Compare-Report ($inputReport, $differenceReport)
{
	0..($differenceReport.Mailbox.Count - 1) | %{$differenceReport.Mailbox[$_] = (Compare-Mailbox -inputMailbox $inputReport.Mailbox[$_] -differenceMailbox $differenceReport.Mailbox[$_])}
	0..($differenceReport.Folder.Count - 1) | %{$differenceReport.Folder[$_] = (Compare-Folder -inputFolder $inputReport.Folder[$_] -differenceFolder $differenceReport.Folder[$_])}
	$differenceReport.Database = (Compare-Database -inputDatabase $inputReport.Database -differenceDatabase $differenceReport.Database)
	
	$differenceReport
}

#################################################/Differencing Functions##############################################

#################################################File Functions#######################################################

####Export-Report##
#[Void]
#Exports a copy of a report (database, mailbox, and folder statistics) to the specified path or current directory in *.XML format
######################

Function Export-Report ($mbxStat, $mbxFolderStat, $dataBase, $reportPath = ".\")
{
	$report = New-Object PSObject -Property @{"Database"=$dataBase; "Mailbox"=$mbxStat; "Folder"=$mbxFolderStat}

	$viablePath = Test-Path $reportPath
	
	If ($reportPath.EndsWith("\") -ne $True)
	{
		$writePath = $reportPath.PadRight(($reportPath.Length + 1), "\")
	}
	
	Else
	{
		$writePath = $reportPath
	}
	
	If ($viablePath -eq $True)
	{
		$report | Select * | Export-CLIXML ($writePath + $dataBase.Name + (Get-TimeStamp) + '.XML')
	}
	
	Else
	{
		$message = ('The specified path ' + '"' + $($reportPath) + '"' + ' is invalid.' + ' Check that the path exists and that you have write access.')
	
		Write-Warning -Message $message
	}
}

####Convert-Report####
#[Void]
#Imports an *.XML report and exports it to *.CSV format.
######################
Function Convert-Report ($report,$reportPath)
{
	$viablePath = Test-Path $reportPath
	
	If ($reportPath.EndsWith(".XML") -ne $True)
	{
		$message = ('The specified path ' + '"' + $($reportPath) + '"' + ' is invalid.' + ' Check that the path exists and that you have write access.')
		
		Write-Warning -Message $message
	}
	
	If ($viablePath -eq $True)
	{
		$report.Database | Select * | Export-CSV $reportPath.Replace(".XML","-Database.CSV")
		$report.Mailbox | Select * | Export-CSV $reportPath.Replace(".XML","-Mailbox.CSV")
		$report.Folder | Select * | Export-CSV $reportPath.Replace(".XML","-Folder.CSV")		
	}
	
	Else
	{
		$message = ('The specified path ' + '"' + $($reportPath) + '"' + ' is invalid.' + ' Check that the path exists and that you have write access.')
	
		Write-Warning -Message $message
	}
}

####Import-Report#####
#[Report]
#Imports the report details from the specified file path (database, mailbox, and folder statistics)
######################

Function Import-Report ($reportPath)
{

	$validPath = Test-Path $reportPath


	If ($validPath -and $reportPath.ToString().EndsWith(".XML"))
	{
		$report = Import-CLIXML $reportPath

	}
	
	Else
	{
		$message = ('The specified path ' + '"' + $reportPath + '"' + ' does not exist, does not reference a valid *.XML file, or you do not have write access.')
	
		Write-Warning $message
	}
	
	$report
	
}

####Out-Report########
#[Report]
#Outputs database details and top 25 mailboxes by size and top 25 folders by size
######################

Function Out-Report ($inputReport, $topItems = 25)
{
	Write-Host ("Database Details:`n")
	$inputReport.Database | FL Name,Guid
	$inputReport.Database | FT @{Label="Item Size";Expression={Convert-Size $_.ItemSize}}, @{Label="Deleted Item Size";Expression={Convert-Size $_.DeletedItemSize}}, @{Label="Size On Disk";Expression={Convert-Size $_.PhysicalSize}}, @{Label="Available New Mailbox Space";Expression={Convert-Size $_.WhiteSpace}}, @{Label="Expected Overhead";Expression={"{0:P2}" -f ($_.ExpectedOverhead / $_.PhysicalSize)}}, @{Label="Actual Overhead";Expression={"{0:P2}" -f ($_.ActualOverhead / $_.PhysicalSize)}} -AutoSize
	
	$inputReport.Database | FT @{Label="Connected Mailbox Size";Expression={Convert-Size $_.ConnectedMailboxSize}}, @{Label="Disconnected Mailbox Size";Expression={Convert-Size $_.DisconnectedMailboxSize}}
	
	
	Write-Host ("Mailbox Details:`n")
	Write-Host ("Top " + $topItems.ToString() + " mailboxes by ItemSize")
	$inputReport.Mailbox | Sort ItemSize -Descending | Select -First $topItems | FT DisplayName,Items,DeletedItems,AssociatedItems,@{Label="ItemSize";Expression={Convert-Size $_.ItemSize}},@{Label="DeletedItemSize";Expression={Convert-Size $_.DeletedItemSize}} -AutoSize
	
	Write-Host ("Top " + $topItems.ToString() + " mailboxes by DeletedItemSize")
	$inputReport.Mailbox | Sort DeletedItemSize -Descending | Select -First $topItems | FT DisplayName,Items,DeletedItems,AssociatedItems,@{Label="ItemSize";Expression={Convert-Size $_.ItemSize}},@{Label="DeletedItemSize";Expression={Convert-Size $_.DeletedItemSize}} -AutoSize
	
	Write-Host ("Top " + $topItems.ToString() + " mailboxes by Items")
	$inputReport.Mailbox | Sort Items -Descending | Select -First $topItems | FT DisplayName,Items,DeletedItems,AssociatedItems,@{Label="ItemSize";Expression={Convert-Size $_.ItemSize}},@{Label="DeletedItemSize";Expression={Convert-Size $_.DeletedItemSize}} -AutoSize
	
	Write-Host ("Top " + $topItems.ToString() + " mailboxes by DeletedItems")
	$inputReport.Mailbox | Sort DeletedItems -Descending | Select -First $topItems | FT DisplayName,Items,DeletedItems,AssociatedItems,@{Label="ItemSize";Expression={Convert-Size $_.ItemSize}},@{Label="DeletedItemSize";Expression={Convert-Size $_.DeletedItemSize}} -AutoSize
	
	Write-Host ("Top " + $topItems.ToString() + " mailboxes by AssociatedItems")
	$inputReport.Mailbox | Sort AssociatedItems -Descending | Select -First $topItems | FT DisplayName,Items,DeletedItems,AssociatedItems,@{Label="ItemSize";Expression={Convert-Size $_.ItemSize}},@{Label="DeletedItemSize";Expression={Convert-Size $_.DeletedItemSize}} -AutoSize
	
	
	Write-Host ("Folder Details:`n")
	Write-Host ("Top " + $topItems.ToString() + " folders by Size")
	$inputReport.Folder | Sort Size -Descending | Select -First $topItems |  FT @{Label="DisplayName";Expression={(Get-Mailbox $_.Identity.ToString().Split("\")[0]).DisplayName}},FolderPath,ItemCount,DeletedItemCount,@{Label="Size";Expression={Convert-Size $_.Size}} -AutoSize
	
	Write-Host ("Top " + $topItems.ToString() + " folders by ItemCount")
	$inputReport.Folder | Sort ItemCount -Descending | Select -First $topItems |  FT @{Label="DisplayName";Expression={(Get-Mailbox $_.Identity.ToString().Split("\")[0]).DisplayName}},FolderPath,ItemCount,DeletedItemCount,@{Label="Size";Expression={Convert-Size $_.Size}} -AutoSize
	
	Write-Host ("Top " + $topItems.ToString() + " folders by DeletedItemCount")
	$inputReport.Folder | Sort DeletedItemCount -Descending | Select -First $topItems |  FT @{Label="DisplayName";Expression={(Get-Mailbox $_.Identity.ToString().Split("\")[0]).DisplayName}},FolderPath,ItemCount,DeletedItemCount,@{Label="Size";Expression={Convert-Size $_.Size}} -AutoSize
}

#################################################Utility Functions###################################################

#DatabaseSpaceDump####
#[Void]
#Collects an ESEUTIL /MS from a passive copy of the specified database and writes to *.TXT
######################

Function DatabaseSpaceDump ($exportPath = ".\")
{
	$validPath = Test-Path $exportPath
	If ($exportPath.EndsWith("\") -ne $True)
	{
		$writePath = $exportPath.PadRight(($exportPath.Length + 1), "\")
	}
	
	Else
	{
		$writePath = $exportPath
	}
	
	#####
	
	$DB_Status = get-mailboxdatabasecopystatus
	
	Write-host " "
	Write-host " "
	write-host -foregroundcolor Green "This Function is designed to get a space dump (eseutil /ms) to assist in determining the amount of true whitespace in a database. Please note that if you select Yes to capture the space dump, the Microsoft Exchange Replication service will be temporarily stopped during the capture. No databases that are currently mounted will be dismounted."

	Write-host " "
	Write-host " "
	write-host -foregroundcolor Yellow "Looking for Passive Database Copies:"
	foreach ($object in $DB_Status)
	{
		$ServerName = gc env:computername
		$DatabaseStatus = Get-MailboxDatabase -Identity $object.DatabaseName -Status
		$DatabaseName = $object.DatabaseName
		$Database = Get-MailboxDatabase -Identity $DatabaseName
		$CheckLagCopyHashTable = $Database.ReplayLagTimes
		$HashEntry = $CheckLagCopyHashTable | ?{$_.key -eq $ServerName}
		$HashValue = $HashEntry.value
		$Days = $HashValue.days
		$Hours = $HashValue.hours
		$Minutes = $HashValue.minutes
		$Seconds = $HashValue.seconds

		
	
		if ($object.status -ne "Mounted")
		{
			if ([int]$Days -ne '0' -OR [int]$Hours -ne '0' -OR [int]$Minutes -ne '0' -OR [int]$Seconds -ne '0')
			{
				Write-host " "
				write-host -foregroundcolor Yellow "!!!WARNING!!! The database copy"$object.DatabaseName"on this server"$ServerName" appears to be a lagged Copy"
				Write-host " "
				write-host -foregroundcolor Yellow "One or more of the Replay Lag Times have been configured" 
				write-host -foregroundcolor Yellow "Replay Lag times are set at "  $Days "Days," $Hours "Hours," $Minutes "Minutes," $Seconds "Seconds"
				write-host -foregroundcolor Yellow "It is not recommend to take a space dump of a lagged copy as doing so will replay all of the lagged transaction logs into the database to bring the database copy into a clean shut down.  Proceed with Caution"
			}

				Write-host -foregroundcolor Green "Would you like to take a Space dump of: "$object.DatabaseName"? Enter 1 for Yes or 2 for no" -nonewline; $Selection = Read-Host " "
				if ($Selection -eq 1)
				{
					$EDBPath = $DatabaseStatus.edbfilepath
					$LogPath = $DatabaseStatus.LogFolderPath
					write-host -foregroundcolor Yellow "Stopping MSExchange Replication Service"
					stop-service "Microsoft Exchange Replication"

					write-host -foregroundcolor yellow "Starting Soft Recovery on Passive Copy."
					$DatabaseStatus | % {eseutil /r $_.LogFilePrefix /L$LogPath /S$LogPath /a}
					write-host -foregroundcolor Yellow "Finished Soft Recovery on Passive Copy"
					
					write-host -foregroundcolor Yellow "Starting Space Dump Passive Copy. Writing File to "$writePath""$DatabaseName"-SpaceDump.TXT."
					If ($validPath)
					{
						$DatabaseStatus | % {eseutil /ms $_.edbfilepath} > ($writePath + $DatabaseName + "-SpaceDump.TXT")
					}
					
					Else
					{
						Write-Warning -Message ("The specified path " + "'" + $writePath + "'" + "does not exist or you do not have write access.")
					}
					write-host -foregroundcolor Yellow "Space Dump Completed, restarting replication service"
					start-service "Microsoft Exchange Replication"
				}
		}
		else
		{
			Write-host $object.DatabaseName "is a mounted database on this server."
		}
	}
}

#DatabaseIssueEvents##
#[Void]
#Searches for events concerning Online Maintenance Overlap and Possible Corruption, outputting them to the screen
######################

Function DatabaseIssueEvents  ($exportPath = ".\")
{
	Write-Host "Function to capture possible OLM Overlap events or Possible DB Corruption events"

	$LoggedEvents = Get-EventLog -LogName Application 

	$ObjArray= @()
	Foreach ($Event in $LoggedEvents)
	{
		If ($Event.EventID -eq 508 -OR $Event.EventID -eq 1114 -OR $Event.EventID -eq 628 -OR $Event.EventID -eq 629 -OR $Event.EventID -eq 531 -OR $Event.EventID -eq 510 -OR $Event.EventID -eq 222 -OR $Event.EventID -eq 509 -OR $Event.EventID -eq 139 -OR $Event.EventID -eq 533 -OR $Event.EventID -eq 203 -OR $Event.EventID -eq 233 -OR $Event.EventID -eq 474 -OR $Event.EventID -eq 218)
		{
			$ObjArray = $ObjArray + $Event			
		}


	}

	If($objArray -ne "" -OR $objArray -ne $null)
	{ 
		write-host "The Following Events might indicate possible DB corruption for the DB they were logged for"

		&{
			Foreach($obj in $ObjArray)
			{
				$obj 

			} 
		} | FT -wrap -autosize eventid,timegenerated,source,message
	}
	else
	{
		write-host "No Corruption Type Events were found"
	}
}

#ExportStoreUsageStats#
#[Void]
#Collects and exports current Store Usage Statistics to *.CSV
#######################

Function ExportStoreUsageStats  ($exportPath = ".\")
{
		$validPath = Test-Path $exportPath
	If ($exportPath.EndsWith("\") -ne $True)
	{
		$writePath = $exportPath.PadRight(($exportPath.Length + 1), "\")
	}
	
	Else
	{
		$writePath = $exportPath
	}
	
	#####

	$DB_Status = get-mailboxdatabasecopystatus
	$ServerName = gc env:computername
	write-host " This function will generate a CSV report for Store Usage Statistics for the server or for a specific database."
	write-host " This report will show the top 25 store resource users per server or per database."
	write-host " This information can be reviewed in conjunction with other data captured to troubleshoot store growth issues."
	write-host " Note this may take a few minutes to complete."
	Write-host -foregroundcolor Green "Would you like to get store usage statistics for server: "$ServerName"? Enter 1 for Yes or 2 for no" -nonewline; $Selection = Read-Host " "
	if ($Selection -eq 1)
	{
		write-host "Exporting Server Store Usage statistics to "$writePath""$ServerName"-StoreUsage.CSV"
		If ($validPath)
		{
			Get-StoreUsageStatistics -Server $ServerName | Select * | Export-CSV ($writePath + $ServerName + "-StoreUsage.CSV")
		}
		
		Else
		{
			Write-Warning -Message ("The specified path " + "'" + $writePath + "'" + "does not exist or you do not have write access.")
		}
		
	}
	Write-host "Checking for Mounted Databases"
	foreach ($object in $DB_Status)
	{
		$DatabaseName = $object.DatabaseName	
		if ($object.status -eq "Mounted")
		{
			Write-host -foregroundcolor Green "Would you like to get store usage statistics for database: "$DatabaseName"? Enter 1 for Yes or 2 for no" -nonewline; $Selection = Read-Host " "
			if ($Selection -eq 1)
			{
				write-host "Exporting Database "$DatabaseName" Store Usage statistics to "$writePath""$DatabaseName"-StoreUsage.CSV."
				If ($validPath)
				{
					Get-StoreUsageStatistics -Database $DatabaseName | Select * | Export-CSV ($writePath + $DatabaseName + "-StoreUsage.CSV")
				}
				
				Else
				{
					Write-Warning -Message ("The specified path " + "'" + $writePath + "'" + "does not exist or you do not have write access.")
				}
			}
		}								
	}
}

#################################################/Utility Functions##################################################

#################################################Test################################################################

#Function Test ([int]$isTest=0)
#{
#	If ($isTest)
#	{
		#$mbxStat = Get-mbxStatistics DB01
		#$mbxFolderStat = Get-mbxFolderStatistics DB01
		#$database = Get-dbStatistics DB01

		#

		#$dbCompare = Compare-Database -inputDatabase $database -differenceDatabase $database
		#$dbCompare | FL PhysicalSize,ItemSize,DeletedItemSize,WhiteSpace,ActualOverhead,ExpectedOverhead #-AutoSize

		#$mbxCompare = Compare-Mailbox -inputMailbox $mbxStat[4] -differenceMailbox $mbxStat[4]
		#$mbxCompare | FL ItemSize,DeletedItemSize,AssociatedItems,Items,DeletedItems #-AutoSize

		#$folderCompare = Compare-Folder -inputFolder $mbxFolderStat[30] -differenceFolder $mbxFolderStat[30]
		#$folderCompare | FL Size,ItemCount,DeletedItemCount #-AutoSize

		#

		#Export-Report -mbxStat $mbxstat -mbxFolderStat $mbxfolderstat -dataBase $database -reportPath c:\script\

		#

		#$testImport = Import-Report -reportPath *.XML
		#$testImport

		#

		#$difference = Compare-Report -inputReport $testImport -differenceReport $testImport
		#$difference

		#

		#DatabaseSpaceDump
		#DatabaseIssueEvents
		#ExportStoreUsageStats
#		}
#}
#################################################/Test###############################################################



#Body################################################################################################################
#####################################################################################################################
Switch ($Mode)
{
	0 {Write-Host ("Options for -Mode parameter: `n0=List Options`n1=Collect and Export Size Details (this gets database, mailbox, and folder statistics).`n2=Review Size Report`n3=Compare Size Reports`n4=Export Store Usage Statistics`n5=List Maintenance and Possible Corruption Events`n6=Export an ESEUTIL /MS space dump`n7=Import an *.XML report and convert it into three separate *.CSVs (Database, Mailbox, Folder)")}
	
	1 {
		Write-Host ("Enter the path where you would like to export the *.XML results.  Default is the current working directory.`n")
		$path = Read-Host
		$path = $path.Replace('"','')
		If ($path -eq "") {$path = ".\"}		
	
		Write-Host ("Please enter the name of the database to run reporting against.")
		$dbSelection = Read-Host
		$dbSelection = $dbSelection.Replace('"','')
		
		$mbxStat = Get-mbxStatistics $dbSelection
		$mbxFolderStat = Get-mbxFolderStatistics $dbSelection
		$dbStat = Get-dbStatistics $dbSelection
		
		$report = New-Object PSObject -Property @{"Database"=$dbStat; "Mailbox"=$mbxStat; "Folder"=$mbxFolderStat}
		
		Export-Report -mbxStat $mbxStat -mbxFolderStat $mbxFolderStat -dataBase $dbStat -reportPath $path
		
		Write-Host ("Currently selecting the top 25 items for reporting. Enter another value or press <Enter> to continue.`n")
		$top = Read-Host
		
		If ($top -eq "") {$top = 25}
		
		Out-Report -inputReport $report -topItems $top
	
	}
	
	2 {
		Write-Host ("Enter the FULL file path to the report you would like to review. `n")
		$path = Read-Host
		
		$path = $path.Replace('"','')
		
		$inputReport = Import-Report -reportPath $path		
		
		Write-Host ("Currently selecting the top 25 items for reporting. If you would like to specify a different amount please input that now.  Otherwise press <Enter> to continue with default.`n")
		$top = Read-Host
		
		If ($top -eq "") {$top = 25}
		
		Out-Report -inputReport $inputReport -topItems $top
		
	}
	
	3 {
		Write-Host ("Please enter the FULL path for the older report.`n")
		$inPath = Read-Host
		$inPath = $inPath.Replace('"','')
		$inputReport = Import-Report $inPath
		
		#
		
		Write-Host ("Please enter the FULL path for the newer report.`n")
		$diffPath = Read-Host
		$diffPath = $diffPath.Replace('"','')
		$diffReport = Import-Report $diffPath		
		
		#
		
		$comparisonReport = Compare-Report -inputReport $inputReport -differenceReport $diffReport
		
		Write-Host ("Currently selecting the top 25 items for reporting. If you would like to specify a different amount please input that now.  Otherwise press <Enter> to continue with default.`n")
		$top = Read-Host
		
		If ($top -eq "") {$top = 25}
		
		Out-Report -inputReport $comparisonReport -topItems $top
		
	}
	
	4 {
		Write-Host ("Enter the folder path where you would like to output the results. Default is the current working directory.`n")
		$path = Read-Host
		$path = $path.Replace('"','')
		If ($path -eq "") {$path = ".\"}	
		
		ExportStoreUsageStats -exportPath $path
	}
	
	5 {
		DatabaseIssueEvents
	}
	
	6 {
		Write-Host ("Enter the folder path where you would like to output the results. Default is the current working directory.`n")
		$path = Read-Host
		$path = $path.Replace('"','')
		If ($path -eq "") {$path = ".\"}	
		
		DatabaseSpaceDump -exportPath $path
	}
	
	7 {
		Write-Host ("Enter the FULL path to the report you would like to convert.`n")
		$path = Read-Host		
		$path = $path.Replace('"','')
		
		$inputReport = Import-Report -reportPath $path	
		
		Convert-Report -report $inputReport -reportPath $path
	
	}
	
#	8 {$Test = 1}
	
	default {Write-Host ("This is not a valid selection (0-6).")}
	
}
#/Body################################################################################################################
#####################################################################################################################