<#
*********************************************************************************************************
*                                                                                                       *
*** This powershell script is used to set maintenance windows based on CSV file                       ***
*                                                                                                       *
*********************************************************************************************************
* Created by Ioan Popovici, 03/07/2016  | Requirements: Powershell 3.0, SCCM client SDK, local FS only  *
* ======================================================================================================*
* Modified by                   | Date       | Version  | Comments                                      *
*_______________________________________________________________________________________________________*
* Ioan Popovici                 | 03/07/2016 | v1.0     | First version                                 *
* Ioan Popovici                 | 04/07/2016 | v2.0     | Vastly improved                               *
* Ioan Popovici                 | 07/07/2016 | v3.0     | Added FileSystemWatcher and workaround        *
* Ioan Popovici                 | 07/07/2016 | v3.1     | Cleanup and run from shell optimisations      *
* Ioan Popovici                 | 12/09/2016 | v3.2     | Added MW type                                 *
* Ioan Popovici                 | 12/09/2016 | v3.3     | Improved logging and variable Naming          *
* Ioan Popovici                 | 12/09/2016 | v3.4     | Overall improvements                          *
*-------------------------------------------------------------------------------------------------------*
* Execute with: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoExit -File                 *
* Set-ClientMaintenanceWindows.ps1                                                                      *
*********************************************************************************************************

    .SYNOPSIS
        This powershell script is used to set maintenance windows based on CSV file.
    .DESCRIPTION
        This powershell script is used to set maintenance windows, triggered when the settings CSV file is saved.
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

    ## Cleaning prompt history (for testing only)
    CLS

    ## Get script path and name
    [string]$ScriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
    [string]$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)

    ## CSV and log file initialization
    #  Set the CSV file name
    [string]$csvFileName = $ScriptName

    #  Get CSV file name with extension
    [string]$csvFileNameWithExtension = $ScriptName+'.csv'

    #  Assemble CSV file path
    [string]$csvFilePath = (Join-Path -Path $ScriptPath -ChildPath $csvFileName)+'.csv'

    #  Assemble log file Path
    [string]$LogFilePath = (Join-Path -Path $ScriptPath -ChildPath $ScriptName)+'.log'

    ## Initialize CSV file read time with current time
	[datetime]$csvFileReadTime = (Get-Date)

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Write-Log
Function Write-Log {
<#
.SYNOPSIS
    Writes data to file log, event log and console.
.DESCRIPTION
    Writes data to file log, event log and console.
.PARAMETER EventLogEntryMessage
    The event log entry message.
.PARAMETER EventLogName
    The event log to write to.
.PARAMETER FileLogName
    The file log name to write to.
.PARAMETER EventLogEntrySource
    The event log entry source.
.PARAMETER EventLogEntryID
    The event log entry ID.
.PARAMETER EventLogEntryType
    The event log entry type (Error | Warning | Information | SuccessAudit | FailureAudit).
.PARAMETER SkipEventLog
    Skip writing to event log.
.EXAMPLE
    Write-Log -EventLogEntryMessage 'Set-ClientMW was successful' -EventLogName 'Configuration Manager' -EventLogEntrySource 'Script' -EventLogEntryID '1' -EventLogEntryType 'Information'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    http://sccm-zone.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [Alias('Message')]
        [string]$EventLogEntryMessage,
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('EName')]
        [string]$EventLogName = 'Configuration Manager',
        [Parameter(Mandatory=$false,Position=2)]
        [Alias('Source')]
        [string]$EventLogEntrySource = $ScriptName,
        [Parameter(Mandatory=$false,Position=3)]
        [Alias('ID')]
        [int32]$EventLogEntryID = 1,
        [Parameter(Mandatory=$false,Position=4)]
        [Alias('Type')]
        [string]$EventLogEntryType = 'Information',
        [Parameter(Mandatory=$false,Position=5)]
        [Alias('SkipEL')]
        [switch]$SkipEventLog
    )

    ## Initialization
    #  Getting the date and time
    [string]$LogTime = (Get-Date -Format 'dd-MM-yyyy HH:mm:ss').ToString()

    #  Archive log file if it exists and it's larger than 50 KB
    If ((Test-Path $LogFilePath) -and (Get-Item $LogFilePath).Length -gt 50KB) {
        Get-ChildItem -Path $LogFilePath | Rename-Item -NewName { $_.Name -Replace '.log','.lo_' } -Force
    }

    #  Create event log and event source if they do not exist
    If (-not ([System.Diagnostics.EventLog]::Exists($EventLogName)) -or (-not ([System.Diagnostics.EventLog]::SourceExists($EventLogEntrySource)))) {

        #  Create new event log and/or source
        New-EventLog -LogName $EventLogName -Source $EventLogEntrySource
    }

    ## Error logging
    If ($_.Exception) {

        #  Write to log
        Write-EventLog -LogName $EventLogName -Source $EventLogEntrySource -EventId $EventLogEntryID -EntryType 'Error' -Message "$EventLogEntryMessage `n$_"

        #  Write to console
        Write-Host `n$EventLogEntryMessage -BackgroundColor Red -ForegroundColor White
        Write-Host $_.Exception -BackgroundColor Red -ForegroundColor White
    }
    Else {

        #  Skip event log if requested
        If ($SkipEventLog) {

            #  Write to console
            Write-Host $EventLogEntryMessage -BackgroundColor White -ForegroundColor Blue
        }
        Else {

            #  Write to event log
            Write-EventLog -LogName $EventLogName -Source $EventLogEntrySource -EventId $EventLogEntryID -EntryType $EventLogEntryType -Message $EventLogEntryMessage

            #  Write to console
            Write-Host $EventLogEntryMessage -BackgroundColor White -ForegroundColor Blue
        }
    }

    ##  Assemble log line
    [string]$LogLine = "$LogTime : $EventLogEntryMessage"

    ## Write to log file
    $LogLine | Out-File -FilePath $LogFilePath -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Continue'
}
#endregion

#region Function Get-MaintenanceWindows
Function Get-MaintenanceWindows {
<#
.SYNOPSIS
    Get existing maintenance windows.
.DESCRIPTION
    Get the existing maintenance windows for a collection.
.PARAMETER CollectionName
    Set the collection name for which to list the maintenance Windows.
.EXAMPLE
    Get-MaintenanceWindows -Collection 'Computer Collection'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    http://sccm-zone.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName
    )

    ## Get CollectionID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction 'Stop').CollectionID
    }

    #  Write to log in case of failure
    Catch {
        Write-Log -Message "Getting $CollectionName ID - Failed!"
    }

    ## Get collection maintenance windows
    Try {
        Get-CMMaintenanceWindow -CollectionId $CollectionID -ErrorAction 'Stop'
    }

    #  Write to log in case of failure
    Catch {
        Write-Log -Message "Get maintenance windows for $CollectionName - Failed!"
    }
}
#endregion

#region Function Remove-MaintenanceWindows
Function Remove-MaintenanceWindows {
<#
.SYNOPSIS
    Remove ALL existing maintenance windows.
.DESCRIPTION
    Remove ALL existing maintenance windows from a collection.
.PARAMETER CollectionName
    The collection name for which to remove the maintenance windows.
.EXAMPLE
    Remove-MaintenanceWindows -Collection 'Computer Collection'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    http://sccm-zone.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName
    )

    ## Get collection ID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction 'Stop').CollectionID
    }
    Catch {
        Write-Log -Message "Getting $CollectionName ID - Failed!"
    }

    ## Get collection maintenance windows and delete them
    Try {
        Get-CMMaintenanceWindow -CollectionId $CollectionID | ForEach-Object {
            Remove-CMMaintenanceWindow -CollectionID $CollectionID -Name $_.Name -Force -ErrorAction 'Stop'
            Write-Log -Message ($_.Name+' - Removed!') -SkipEventLog
        }
    }
    Catch {

        #  Write to log in case of failure
        Write-Log -Message "$_.Name  - Removal Failed!"
    }
}
#endregion

#region Function Set-MaintenanceWindows
Function Set-MaintenanceWindows {
<#
.SYNOPSIS
    Set maintenance windows.
.DESCRIPTION
    Set maintenance windows to a collection.
.PARAMETER CollectionName
    The collection name for which to set maintenance windows.
.PARAMETER Date
    The maintenance window date.
.PARAMETER StartTime
    The maintenance window start time.
.PARAMETER StopTime
    The maintenance window stop time.
.PARAMETER ApplyTo
    Maintenance window applicability (Any | SoftwareUpdates | TaskSequences).
.EXAMPLE
    Set-MaintenanceWindows -CollectionName 'Computer Collection' -Date '01/09/2017' -StartTime '01:00'  -StopTime '02:00' -ApplyTo SoftwareUpdates
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    http://sccm-zone.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName,
        [Parameter(Mandatory=$true,Position=1)]
        [Alias('Da')]
        [string]$Date,
        [Parameter(Mandatory=$true,Position=2)]
        [Alias('SartT')]
        [string]$StartTime,
        [Parameter(Mandatory=$true,Position=3)]
        [Alias('StopT')]
        [string]$StopTime,
        [Parameter(Mandatory=$true,Position=4)]
        [Alias('Apply')]
        [string]$ApplyTo
    )

    ## Get collection ID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction 'Stop').CollectionID
    }
    Catch {

        #  Write to log in case of failure
        Write-Log -Message "Getting $CollectionName ID - Failed!"
    }

    ## Setting maintenance window start and stop times
    $MWStartTime = Get-Date -Format 'dd/MM/yyyy HH:mm' -Date ($Date+' '+$StartTime)
    $MWStopTime = Get-Date -Format 'dd/MM/yyyy HH:mm' -Date ($Date+' '+$StopTime)

    ## Create the schedule token
    $MWSchedule = New-CMSchedule -Start $MWStartTime -End $MWStopTime -NonRecurring

    ## Set maintenance window naming convention
    If ($ApplyTo -eq 'Any') { $MWType = 'MWA' }
    ElseIf ($ApplyTo -match 'Software') { $MWType = 'MWU' }
    ElseIf ($ApplyTo -match 'Task') { $MWType = 'MWT' }

    # Set maintenance window name
    $MWName =  $MWType+'.NR.'+(Get-Date -Uformat %Y.%B.%d $MWStartTime)+'.'+$StartTime+'-'+$StopTime

    ## Set maintenance window on collection
    Try {
        $SetNewMW = New-CMMaintenanceWindow -CollectionID $CollectionID -Schedule $MWSchedule -Name $MWName -ApplyTo $ApplyTo -ErrorAction 'Stop'

        #  Write to log
        Write-Log -Message "$MWName - Set!" -SkipEventLog
    }
    Catch {

        #  Write to log in case of failure
        Write-Log -Message "Setting $MWName on $CollectionName - Failed!"
    }
}
#endregion

#region  Function Send-Mail
Function Send-Mail {
<#
.SYNOPSIS
    Send E-Mail to specified address.
.DESCRIPTION
    Send E-Mail body to specified address.
.PARAMETER From
    Source.
.PARAMETER To
    Destination.
.PARAMETER CC
    Carbon copy.
.PARAMETER Body
    E-Mail body.
.PARAMETER SMTPServer
    E-Mail SMTPServer.
.PARAMETER SMTPPort
    E-Mail SMTPPort.
.EXAMPLE
    Set-Mail -Body 'Test' -CC 'test@visma.com'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    http://sccm-zone.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [string]$From = 'SCCM Site Server <noreply@visma.com>',
        [Parameter(Mandatory=$false,Position=1)]
        [string]$To = 'SCCM Team <SCCM-Team@visma.com>',
        [Parameter(Mandatory=$false,Position=2)]
        [string]$CC,
        [Parameter(Mandatory=$false,Position=3)]
        [string]$Subject = 'Info: Maintenance Window Set!',
        [Parameter(Mandatory=$true,Position=4)]
        [string]$Body,
        [Parameter(Mandatory=$false,Position=5)]
        [string]$SMTPServer = 'mail.datakraftverk.no',
        [Parameter(Mandatory=$false,Position=6)]
        [string]$SMTPPort = "25"
    )

    Try {
        If ($CC) {
            Send-MailMessage -From $From -To $To -Subject $Subject -CC $CC -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -ErrorAction 'Stop'
        }
        Else {
            Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -ErrorAction 'Stop'
        }
    }
    Catch {
        Write-Log -Message 'Send Mail - Failed!'
    }
}
#endregion

#region Function Start-DataProcessing
Function Start-DataProcessing {
<#
.SYNOPSIS
    Used for main data processing.
.DESCRIPTION
    Used for main data processing, for this script only.
.EXAMPLE
    Start-DataProcessing
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    http://sccm-zone.com
#>
    ## Import SCCM PSH module and changing context
    Try {
        Import-Module $env:SMS_ADMIN_UI_PATH.Replace('\bin\i386','\bin\configurationmanager.psd1') -ErrorAction 'Stop'
    }
    Catch {
        Write-Log -Message 'Importing SCCM PSH module - Failed!'
    }

    #  Get the CMSITE SiteCode and change connection context
    $SiteCode = Get-PSDrive -PSProvider CMSITE

    #  Change the connection context
    Set-Location "$($SiteCode.Name):\"

    ##  Import the CSV file
    Try {
        $csvFileData = Import-Csv -Path $csvFilePath -Encoding 'UTF8' -ErrorAction 'Stop'
    }
    Catch {
        Write-Log -Message 'Importing CSV Data - Failed!'
    }

    ## Process imported CSV file data
    $csvFileData | ForEach-Object {

        #  Check if we need to remove existing maintenance windows
        If ($_.RemoveExisting -eq 'YES' ) {

            #  Write to log
            Write-Log -Message ('Removing maintenance windows from:  '+$_.CollectionName) -SkipEventLog

            #  Remove maintenance window
            Remove-MaintenanceWindows -CollectionName $_.CollectionName
        }

        #  Set Maintenance Window
        Set-MaintenanceWindows -CollectionName $_.CollectionName -Date $_.Date -StartTime $_.StartTime -StopTime $_.StopTime -ApplyTo $_.ApplyTo
    }

    ## Get maintenance windows for unique collections
    #  Initialize result array
    [array]$Result =@()

    #  Parsing CSV collection names
    $csvFileData.CollectionName | Select-Object -Unique | ForEach-Object {

        #  Getting maintenance windows for collection (split to new line)
        $MaintenanceWindows = Get-MaintenanceWindows -CollectionName $_ | ForEach-Object { $_.Name+"`n" }

        #  Assemble result with descriptors
        $Result+= "`n Listing all maintenance windows for: "+$_+" "+"`n "+$MaintenanceWindows
    }

    #  Convert the result to string and write it to log
    [string]$ResultString = Out-String -InputObject $Result
    Write-Log -Message $ResultString

    ## E-Mail result
    Send-Mail -Body $ResultString

    ## Return to Script Path
    Set-Location $ScriptPath

    ## Remove SCCM PSH Module
    Remove-Module 'ConfigurationManager' -Force -ErrorAction 'Continue'
}

#endregion

#region Function Test-FileChangeEvent
Function Test-FileChangeEvent {
<#
.SYNOPSIS
    Workaround for FileSystemWatcher firing multiple events during a write operation.
.DESCRIPTION
    FileSystemWatcher may fire multiple events on a write operation.
    It's a known problem but it's not a bug in FileSystemWatcher.
    This function is discarding events fired more than once a second.
.PARAMETER $FileLastReadTime
    Specify file last read time, used to avoid global parameter.
.EXAMPLE
    Test-FileChangeEvent -FileLastReadTime $FileLastReadTime
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    http://sccm-zone.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [Alias('ReadTime')]
        [datetime]$FileLastReadTime
    )

    ## Get CSV file last write time
    [datetime]$csvFileLastWriteTime = (Get-ItemProperty -Path $csvFilePath).LastWriteTime

    ## Test if the file change event is valid by comparing the CSV last write time and parameter apecified time
    If (($csvFileLastWriteTime - $FileLastReadTime).Seconds -ge 1) {

        ## Write to log
        Write-Log -Message "`nCSV file change - Detected!" -SkipEventLog

        ## Start main data processing and wait for it to finish
        Start-DataProcessing | Out-Null
    }
    Else {

        ## Do nothing, the file change event was fired more than once a second
    }
}

#endregion

#endregion
##*=============================================
##* END FUNCTION LISTINGS
##*=============================================

##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody

    ## Initialize file qatcher and wait for file changes
    $FileWatcher = New-Object System.IO.FileSystemWatcher
    $FileWatcher.Path = $ScriptPath
    $FileWatcher.Filter = $csvFileNameWithExtension
    $FileWatcher.IncludeSubdirectories = $false
    $FileWatcher.NotifyFilter = [System.IO.NotifyFilters]::'LastWrite'
    $FileWatcher.EnableRaisingEvents = $true

    #  Register file watcher event
    Register-ObjectEvent -InputObject $FileWatcher -EventName 'Changed' -Action {

        # Test if we really need to start processing
        Test-FileChangeEvent -FileLastReadTime $csvFileReadTime

        #  Reinitialize DateTime variable to be used on next file change event
        $csvFileReadTime = (Get-Date)
    }

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
