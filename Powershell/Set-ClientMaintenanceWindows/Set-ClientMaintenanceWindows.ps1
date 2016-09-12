<#
*********************************************************************************************************
*                                                                                                       *
*** This Powershell Script is used to set Maintenance Windows based on CSV File                       ***
*                                                                                                       *
*********************************************************************************************************
* Created by Ioan Popovici, 03/07/2016  | Requirements Powershell 3.0, SCCM Client SDK, Local FS Only   *
* ======================================================================================================*
* Modified by                   | Date       |    Version  | Comments                                   *
*_______________________________________________________________________________________________________*
* Ioan Popovici                 | 03/07/2016 | v1.0     | First version                                 *
* Ioan Popovici                 | 04/07/2016 | v2.0     | Vastly Improved                               *
* Ioan Popovici                 | 07/07/2016 | v3.0     | Added FileSystemWatcher and Workaround        *
* Ioan Popovici                 | 07/07/2016 | v3.1     | Cleanup and Run from Shell Optimisations      *
*-------------------------------------------------------------------------------------------------------*
*                                                                                                       *
*********************************************************************************************************

    .SYNOPSIS
        This Powershell Script is used to set Maintenance Windows based on CSV File.
    .DESCRIPTION
        This Powershell Script is used to set Maintenance Windows based on a CSV File.
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

    ## Run with this Line: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoExit -File E:\SCCM_Client_Scripts\VSW_AccountView\Set-ClientMaintenanceWindows\Set-ClientMaintenanceWindows.ps1

    ## Get Script Path and Name
    [string]$ScriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
    [string]$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)

    ## CSV File Initialization
    #  Set the CSV File Name
    [string]$csvFileName = $ScriptName

    #  Get CSV File Name With Extension
    [string]$csvFileNameWithExtension = $ScriptName+".csv"

    #  Assemble CSV File Path
    [string]$csvFilePath = (Join-Path -Path $ScriptPath -ChildPath $csvFileName)+".csv"

    #  Initialize CSV File Read Time to Current Time
    [datetime]$csvFileReadTime = (Get-Date)

    ## Assemble Log File Path
    [string]$LogFilePath = (Join-Path -Path $ScriptPath -ChildPath $ScriptName)+".log"

    ## Cleaning prompt history
    CLS

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
    Writes an event to EventLog.
.DESCRIPTION
    Writes an event to EventLog with a specified source.
.PARAMETER EventLogName
    The EventLog to write to.
.PARAMETER EventLogEntrySource
    The EventLog Entry Source.
.PARAMETER EventLogEntryID
    The EventLog Entry ID.
.PARAMETER EventLogEntryType
    The EventLog Entry Type. (Error | Warning | Information | SuccessAudit | FailureAudit)
.PARAMETER EventLogEntryMessage
    The EventLog Entry Message.
.EXAMPLE
    Write-Log -EventLogName "Configuration Manager" -EventLogEntrySource "Script" -EventLogEntryID "1" -EventLogEntryType "Information" -EventLogEntryMessage "Set-ClientMW was succesfull"
.NOTES
.LINK
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [Alias('EName')]
        [string]$EventLogName = "Configuration Manager",
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('LName')]
        [string]$LogFileName = $ScriptName,
        [Parameter(Mandatory=$false,Position=2)]
        [Alias('Source')]
        [string]$EventLogEntrySource = $ScriptName,
        [Parameter(Mandatory=$false,Position=3)]
        [Alias('ID')]
        [int32]$EventLogEntryID = 1,
        [Parameter(Mandatory=$false,Position=4)]
        [Alias('Type')]
        [string]$EventLogEntryType = "Information",
        [Parameter(Mandatory=$false,Position=5)]
        [Alias('Message')]
        [string]$EventLogEntryMessage,
        [Parameter(Mandatory=$false,Position=6)]
        [Alias('SkipEL')]
        [switch]$SkipEventLog
    )

    ## Initialization

    #  Getting the Date and time
    [string]$LogTime = (Get-Date -Format 'MM-dd-yyyy HH:mm:ss').ToString()

    #  Archive FileLog if it exists and it's larger than 50 KB
    If ((Test-Path $LogFilePath) -and (Get-Item $LogFilePath).Length -gt 50KB) {
        Get-ChildItem -Path $LogFilePath | Rename-Item -NewName { $_.Name -Replace '\.log','.lo_' } -Force
    }

    # Create EventLog and Event Source if they do not exist
    If (-not ([System.Diagnostics.EventLog]::Exists($EventLogName)) -or (-not ([System.Diagnostics.EventLog]::SourceExists($EventLogEntrySource)))) {

        #  Create new EventLog and/or Source
        New-EventLog -LogName $EventLogName -Source $EventLogEntrySource
    }

    ## Error Logging
    #  If Exception was Triggered
    If($_.Exception) {

        #  Write to EventLog
        Write-EventLog -LogName $EventLogName -Source $EventLogEntrySource -EventId $EventLogEntryID -EntryType "Error" -Message "$EventLogEntryMessage `n$_"

        #  Write to Console
        Write-Host `n$EventLogEntryMessage -BackgroundColor Red -ForegroundColor White
        Write-Host $_.Exception -BackgroundColor Red -ForegroundColor White
    }
    Else {

        #  Skip Event Log if requested
        If ($SkipEventLog) {

            #  Write to Console
            Write-Host $EventLogEntryMessage -BackgroundColor White -ForegroundColor Blue
        }
        Else {

            #  Write to EventLog
            Write-EventLog -LogName $EventLogName -Source $EventLogEntrySource -EventId $EventLogEntryID -EntryType $EventLogEntryType -Message $EventLogEntryMessage

            #  Write to Console
            Write-Host $EventLogEntryMessage -BackgroundColor White -ForegroundColor Blue
        }
    }

    ##  Construct LogLine
    [string]$LogLine = "$LogTime : $EventLogEntryMessage"

    ## Write to Log File
    $LogLine | Out-File -FilePath $LogFilePath -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Stop'
}
#endregion

#region Function Get-MaintenanceWindows
Function Get-MaintenanceWindows {
<#
.SYNOPSIS
    Get Existing Maintenance Windows.
.DESCRIPTION
    Get the existing maintenance windows for a collection.
.PARAMETER CollectionName
    The collection name for which to list the Mainteance Windows.
.EXAMPLE
    Get-MaintenanceWindows -Collection "Computer Collection"
.NOTES
.LINK
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName
    )

    ## Get CollectionID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop -ErrorVariable Error).CollectionID
    }

    #  Write to log in case of failure
    Catch {
        Write-Log -Message "Getting $CollectionName ID Failed!"
    }

    ## Get Collection Maintenance Windows
    Try {
        Get-CMMaintenanceWindow -CollectionId $CollectionID
    }

    #  Write to log in case of failure
    Catch {
        Write-Log -Message "Get Maintenance Windows for $CollectionName - Failed!"
    }
}
#endregion

#region Function Remove-MaintenanceWindows
Function Remove-MaintenanceWindows {
<#
.SYNOPSIS
    Remove ALL Existing Maintenance Windows.
.DESCRIPTION
    Remove ALL Existing Maintenance Windows from a Collection.
.PARAMETER CollectionName
    The collection name for which to remove the Mainteance Windows.
.EXAMPLE
    Remove-MaintenanceWindows -Collection "Computer Collection"
.NOTES
.LINK
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName
    )

    ## Get CollectionID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop -ErrorVariable Error).CollectionID
    }
    Catch {
        Write-Log -Message "Getting $CollectionName ID Failed!"
    }

    ## Get Collection Maintenance Windows and delete them
    Try {
        Get-CMMaintenanceWindow -CollectionId $CollectionID | ForEach-Object {
            Remove-CMMaintenanceWindow -CollectionID $CollectionID -Name $_.Name -Force -ErrorAction Stop -ErrorVariable Error
            Write-Log -Message ($_.Name+" - Removed!") -SkipEventLog
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
    Remove Existing Maintenance Windows.
.DESCRIPTION
    Remove existing maintenance windows from a collection.
.PARAMETER CollectionName
    The collection name for which to remove the Mainteance Windows.
.PARAMETER MWDate
    The Maintenance Window Date.
.PARAMETER MWStartTime
    The Maintenance Window Start Time.
.PARAMETER MWStopTime
    The Maintenance Window Stop Time.
.EXAMPLE
    Set-MaintenanceWindows -CollectionName "Computer Collection" -MWDate "01/09/2017" -MWStartTime "01:00:00"  -MWStopTime "02:00:00"
.NOTES
.LINK
#>
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName,
        [Parameter(Mandatory=$true,Position=1)]
        [Alias('SDate')]
        [string]$StartDate,
        [Parameter(Mandatory=$true,Position=2)]
        [Alias('SartT')]
        [string]$StartTime,
        [Parameter(Mandatory=$true,Position=3)]
        [Alias('StopT')]
        [string]$StopTime
    )

    ## Get CollectionID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop -ErrorVariable Error).CollectionID
    }
    Catch {

        #  Write to log in case of failure
        Write-Log -Message "Getting $CollectionName ID - Failed!"
    }

    $Message = "Setting Maintenance Windows on "+$_.CollectionName+" for "+(Get-Date -Uformat %B $StartDate)
    Write-Log -Message $Message -SkipEventLog

    ## Setting Maintenance Window Start and Stop times
    $MWStartTime = Get-Date -Format "M/dd/yyyy HH:mm:ss" -Date ($StartDate+' '+$StartTime)
    $MWStopTime = Get-Date -Format "M/dd/yyyy HH:mm:ss" -Date ($StartDate+' '+$StopTime)

    ## Create The Schedule Token
    $MWSchedule = New-CMSchedule -Start $MWStartTime -End $MWStopTime -NonRecurring

    ## Set Maintenance Window Naming Convention MW Month
    $MWName =  'MW.NR.'+(Get-Date -Uformat %Y.%B.%d $MWStartDate)+'.'+$StartTime+'-'+$StopTime

    ## Set Maintenance Window
    Try {
        $SetNewMW = New-CMMaintenanceWindow -CollectionID $CollectionID -Schedule $MWSchedule -Name $MWName -ApplyTo Any -ErrorVariable Error -ErrorAction Stop
        Write-Log -Message "Setting $MWName on $CollectionName - Successful" -SkipEventLog
    }
    Catch {

        #  Write to Log and Console in case of failure
        Write-Log -Message "Setting $MWName on $CollectionName - Failed!"
    }
}
#endregion

#region  Function Send-Mail
Function Send-Mail {
<#
.SYNOPSIS
    Send E-Mail to Specified Address.
.DESCRIPTION
    Send E-Mail body to Specified Address.
.PARAMETER From
    Source.
.PARAMETER To
    Destination.
.PARAMETER CC
    Carbon Copy.
.PARAMETER Body
    E-Mail Body.
.PARAMETER SMTPServer
    E-Mail SMTPServer.
.PARAMETER $SMTPPort
    E-Mail SMTPPort.
.EXAMPLE
    Set-Mail -Body "Test" -CC "test@visma.com"
.NOTES
.LINK
#>
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [string]$From = "SCCM Site Server <noreply@visma.com>",
        [Parameter(Mandatory=$false,Position=1)]
        [string]$To = "SCCM Team <SCCM-Team@visma.com>",
        [Parameter(Mandatory=$false,Position=2)]
        [string]$CC,
        [Parameter(Mandatory=$false,Position=3)]
        [string]$Subject = "Info: Maintenance Window Set!",
        [Parameter(Mandatory=$true,Position=4)]
        [string]$Body,
        [Parameter(Mandatory=$false,Position=5)]
        [string]$SMTPServer = "mail.datakraftverk.no",
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
        Write-Log -Message "Send Mail Failed!"
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
.LINK
#>
    ## Import SCCM SDK Module
    Import-Module "E:\SCCM\AdminConsole\bin\ConfigurationManager.psd1" -ErrorAction 'Stop'

    ## Change Path to CM Site
    CD VSM:

    ## Write to FileLog
    Write-Log -Message "Processing Started..." -SkipEventLog

    ## Import the CSV File
    $csvFileData = Import-Csv -Path $csvFilePath -Encoding 'UTF8' #-ErrorAction 'Stop'

    ## Process Imported CSV File Data
    $csvFileData | ForEach-Object {

        #  Check if we need the Remove Existing Maintenance Window is set
        If ($_.RemoveALLExistingMW -eq 'YES' ) {

            #  Write to Log and Console
            $Message = "Remove Existing Maintenance Windows from "+$_.CollectionName
            Write-Log -Message $Message -SkipEventLog

            #  Remove Maintenance Window
            Remove-MaintenanceWindows $_.CollectionName
        }

        #  Set Maintenance Window
        Set-MaintenanceWindows -CollectionName $_.CollectionName -StartDate $_.StartDate -StartTime $_.StartTime -StopTime $_.StopTime
        }

        ## Get Maintanance Windows for Unique Collections

        #  Result variable
        [array]$Result =@()

        #  Parsing CSV Collection Names
        $csvFileData.CollectionName | Select-Object -Unique | ForEach-Object {

            #  Getting Maintenance Windows for Collection (Split to New Line)
            $MaintenanceWindows = Get-MaintenanceWindows -CollectionName $_ | ForEach-Object { $_.Name+"`n" }

            #  Creating Result with Descriptors
            $Result+= "`n Listing All Maintenance Windows for Collection: "+$_+" "+"`n "+$MaintenanceWindows
        }

        #  Convert the Result to string and Write it to the Log, EventLog and Console
        [string]$ResultString = Out-String -InputObject $Result
        Write-Log -Message $ResultString

        ## E-Mail Result
        Send-Mail -Body $ResultString

        ## Write to FileLog
        Write-Log -Message "Processing Finished..." -SkipEventLog

        ## Change Path ScriptRoot
        CD $ScriptPath

        ## Remove SCCM SDK Module
        Remove-Module "ConfigurationManager" -Force -ErrorAction 'Continue'
}

#endregion

#region Function Test-FileChangeEvent
Function Test-FileChangeEvent {
<#
.SYNOPSIS
    Workaround for FileSystemWatcher firing multiple events during a write operation.
.DESCRIPTION
    FileSystemWatcher may fire multiple events on a write operation.
    It's a Known Problem but it's not a Bug in FileSystemWatcher.
    This function is discarding events fired more than once a second, used in this script only.
.PARAMETER $FileLastReadTime
    Specify File Last Read Time, used to avoid global parameter.
.EXAMPLE
    Test-FileChangeEvent -FileLastReadTime $FileLastReadTime
.NOTES
.LINK
#>
    Param (
            [Parameter(Mandatory=$true)]
            [Alias('ReadTime')]
            [datetime]$FileLastReadTime
    )

    ## Get CSV File Last Write Time
    [datetime]$csvFileLastWriteTime = (Get-ItemProperty -Path $csvFilePath).LastWriteTime

    ## Test if the File Change Event is valid by comparing the csv Last Write Time and Parameter Specified Time
    If (($csvFileLastWriteTime - $FileLastReadTime).Seconds -ge 1) {

        ## Write to File Log
        Write-Host ""
        Write-Log -Message "CSV File Change Detected!" -SkipEventLog

        ## Start Main Data Processing and wait for it to finish
        Start-DataProcessing | Out-Null
    }
    Else {

        ## Do Nothing, the File Change Event was fired more than once a second
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

    ## Initialize File Watcher and wait for File Changes
    $FileWatcher = New-Object System.IO.FileSystemWatcher
    $FileWatcher.Path = $ScriptPath
    $FileWatcher.Filter = $csvFileNameWithExtension
    $FileWatcher.IncludeSubdirectories = $false
    $FileWatcher.NotifyFilter = [System.IO.NotifyFilters]::'LastWrite'
    $FileWatcher.EnableRaisingEvents = $true

    write-host "ScriptPath:" $ScriptPath
    write-host "CSV File Name:" $csvFileNameWithExtension

    #  Register File Watcher Event
    Register-ObjectEvent -InputObject $FileWatcher -EventName "Changed" -Action {

        # Test if we really need to Start Processing
        Test-FileChangeEvent -FileLastReadTime $csvFileReadTime

        #  Reinitialize DataTime variable to be used on next File Change Event
        $csvFileReadTime = (Get-Date)
    }

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
