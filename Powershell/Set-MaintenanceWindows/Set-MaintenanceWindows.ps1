<#
*********************************************************************************************************
*                                                                                                       *
*** This Powershell Script is used to set Maintenance Windows based on PatchTuesday on SU Collections ***
*                                                                                                       *
*********************************************************************************************************
* Created by Ioan Popovici, 30/03/2015       | Requirements Powershell 2.0                              *
* ======================================================================================================*
* Modified by                   |    Date    | Revision | Comments                                      *
*_______________________________________________________________________________________________________*
* Ioan Popovici/Octavian Cordos | 30/03/2015 | v1.0     | First version                                 *
* Ioan Popovici/Octavian Cordos | 31/03/2015 | v2.0     | Vastly Improved                               *
* Ioan Popovici                 | 08/01/2016 | v2.1     | Fixed Locale                                  *
* Octavian Cordos               | 11/01/2016 | v2.2     | Improved MW Naming                            *
* Ioan Popovici                 | 11/01/2016 | v2.3     | Added Logging and Error Detection, cleanup    *
* Ioan Popovici                 | 12/09/2016 | v2.4     | Added MW Type                                 *
* Ioan Popovici                 | 12/09/2016 | v2.5     | Improved Logging and Variable Naminc          *
*-------------------------------------------------------------------------------------------------------*
*                                                                                                       *
*********************************************************************************************************

    .SYNOPSIS
        This Powershell Script is used to set Maintenance Windows based on PatchTuesday on SU Collections.
    .DESCRIPTION
        This Powershell Script is used to set Maintenance Windows based on PatchTuesday on SU Collections.
#>

##*=============================================
##* INITIALIZATION
##*=============================================
#region Initialization

## Cleaning prompt history (for testing only)
CLS

## Get Script Path and Name
[string]$ScriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
[string]$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)

## CSV and Log file initialization
#  Set the CSV file name
[string]$csvFileName = $ScriptName

#  Get CSV file name with extension
[string]$csvFileNameWithExtension = $ScriptName+".csv"

#  Assemble CSV file path
[string]$csvFilePath = (Join-Path -Path $ScriptPath -ChildPath $csvFileName)+".csv"

## Assemble Log file Path
[string]$LogFilePath = (Join-Path -Path $ScriptPath -ChildPath $ScriptName)+".log"

## Import modules and files
## Import the CSV file
$csvFileData = Import-Csv -Path $csvFilePath -Encoding 'UTF8'
Import-Module $env:SMS_ADMIN_UI_PATH.Replace("\bin\i386","\bin\configurationmanager.psd1")

## Get the CMSITE SiteCode and change connection context
$SiteCode = Get-PSDrive -PSProvider CMSITE

#  Change the connection context
Set-Location "$($SiteCode.Name):\"

#endregion
##*=============================================
##* END INITIALIZATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Write-Log
Function Write-Log {
<#
.SYNOPSIS
    Writes data to Log, EventLog and Console.
.DESCRIPTION
    Writes data to Log, EventLog and Console.
.PARAMETER EventLogName
    The EventLog to write to.
.PARAMETER FileLogName
    The File Log Name to write to.
.PARAMETER EventLogEntrySource
    The EventLog Entry Source.
.PARAMETER EventLogEntryID
    The EventLog Entry ID.
.PARAMETER EventLogEntryType
    The EventLog Entry Type. (Error | Warning | Information | SuccessAudit | FailureAudit)
.PARAMETER EventLogEntryMessage
    The EventLog Entry Message.
.EXAMPLE
    Write-Log -EventLogEntryMessage "Set-ClientMW was successful" -EventLogName "Configuration Manager" -EventLogEntrySource "Script" -EventLogEntryID "1" -EventLogEntryType "Information"
.NOTES
.LINK
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [Alias('Message')]
        [string]$EventLogEntryMessage,
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('EName')]
        [string]$EventLogName = "Configuration Manager",
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
    If ([System.Diagnostics.EventLog]::Exists($EventLogName) -eq $False -or [System.Diagnostics.EventLog]::SourceExists($EventLogEntrySource) -eq $False) {

        ([System.Diagnostics.EventLog]::Exists($EventLogName) -eq $False -or [System.Diagnostics.EventLog]::SourceExists($EventLogEntrySource) -eq $False)
        #  Create new EventLog and/or Source
        New-EventLog -LogName $EventLogName -Source $EventLogEntrySource
    }

    ## Error Logging
    #  If Exception was Triggered
    If ($_.Exception) {

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
}#endregion

#region Function Get-PatchTuesday
Function Get-PatchTuesday {
<#
.SYNOPSIS
    Calculate Microsoft Patch Tuesday and return it to the pipeline.
.DESCRIPTION
    Get Microsoft Patch tuesday for a specific month.
.PARAMETER Year
    The year for which to calculate Patch Tuesday.
.PARAMETER Month
    The month for which to calculate Patch Tuesday.
.EXAMPLE
    Get-PatchTuesday -Year 2015 -Month 3
.NOTES
.LINK
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Yr')]
        [string]$Year,
        [Parameter(Mandatory=$true,Position=1)]
        [Alias('Mo')]
        [string]$Month
    )

    ## Build Target Month
    [DateTime]$StartingMonth = $Month+"/1/"+$Year

    ## Search for First Tuesday
    While ($StartingMonth.DayofWeek -ine "Tuesday") {
        $StartingMonth = $StartingMonth.AddDays(1)
    }

    ## Set Second Tuesday of the month by adding 7 days
    $PatchTuesday = $StartingMonth.AddDays(7)

    ## Return Patch Tuesday
    Return $PatchTuesday
 }
#endregion

#region Function Remove-MaintenanceWindows
Function Remove-MaintenanceWindows {
<#
.SYNOPSIS
    Remove Existing Maintenance Windows.
.DESCRIPTION
    Remove existing maintenance windows from a collection.
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

            #  Write to Log and Console
            $Message = $_.Name+" - Removed!"
            Write-Log -Message $Message -SkipEventLog
        }
    }
    Catch {

        #  Write to Log and Console in case of failure
        $Message = $_.Name+" - Removal Failed!"
        Write-Log -Message $Message -SkipEventLog
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
.PARAMETER Year
    The Maintenance Window Start Year.
.PARAMETER Month
    The Maintenance Window Start Month.
.PARAMETER OffsetWeeks
    The Maintenance Window offset number of weeks after Patch Tuesday.
.PARAMETER OffsetWeeks
    The Maintenance Window offset number of days after Tuesday.
.PARAMETER StartTime
    The Maintenance Window Start Time.
.PARAMETER StopTime
    The Maintenance Window Stop Time.
.PARAMETER ApplyTo
    Maintenance Window Applys to Any/SoftwareUpdates/TaskSequences.
.EXAMPLE
    Set-MaintenanceWindows -CollectionName "Computer Collection" -Year 2015 -Month 3 -OffsetWeeks 3 -OffsetDays 2 -StartTime "01:00:00"  -StopTime "02:00:00" -ApplyTo SoftwareUpdates
.NOTES
.LINK
#>
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName,
        [Parameter(Mandatory=$true,Position=1)]
        [Alias('Yr')]
        [int16]$Year,
        [Parameter(Mandatory=$true,Position=2)]
        [Alias('Mo')]
        [int16]$Month,
        [Parameter(Mandatory=$true,Position=3)]
        [Alias('Weeks')]
        [int16]$OffsetWeeks,
        [Parameter(Mandatory=$true,Position=4)]
        [Alias('Days')]
        [int16]$OffsetDays,
        [Parameter(Mandatory=$true,Position=5)]
        [Alias('Start')]
        [string]$StartTime,
        [Parameter(Mandatory=$true,Position=6)]
        [Alias('Stop')]
        [string]$StopTime,
        [Parameter(Mandatory=$true,Position=7)]
        [Alias('Apply')]
        [string]$ApplyTo
    )

    ## Get CollectionID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop -ErrorVariable Error).CollectionID
    }
    Catch {

        #  Write to log in case of failure
        Write-Log -Message "Getting $CollectionName ID - Failed!"
    }

    ## Get PatchTuesday
    $PatchTuesday = Get-PatchTuesday $Year $Month

    ## Setting Patch Day, adding offset days and weeks. Get-Date is used to get the date in the same cast, otherwise we cannot convert the date from string to datetime format.
    $PatchDay = (Get-Date -Date $PatchTuesday).AddDays($OffsetDays+($OffsetWeeks*7))

    ## Check if we got ourselves in the next year and return to the main script if true
    If ($PatchDay.Year -gt $Year) {
        Write-Log -Message "Year threshold detected! Ending Cycle..." -SkipEventLog
        Return
    }

    ## Setting Maintenance Window Start and Stop times
    $MWStartTime = (Get-Date -Format "dd/M/yyyy HH:mm" -Date $PatchDay) -Replace "00:00", $StartTime
    $MWStopTime = (Get-Date -Format "dd/M/yyyy HH:mm" -Date $PatchDay) -Replace "00:00", $StopTime

    ## Create The Schedule Token
    $MWSchedule = New-CMSchedule -Start $MWStartTime -End $MWStopTime -NonRecurring

    ## Set Maintenance Window Naming Convention for MW
    If ($ApplyTo -eq "Any") { $MWType = "MWA" }
    ElseIf ($ApplyTo -match "Software") { $MWType = "MWU" }
    ElseIf ($ApplyTo -match "Task") { $MWType = "MWT" }

    $MWName =  $MWType+".NR."+(Get-Date -Uformat %Y.%B.%d $MWStartTime)+"."+$StartTime+"-"+$StopTime

    ## Set Maintenance Window
    Try {
        $SetNewMW = New-CMMaintenanceWindow -CollectionID $CollectionID -Schedule $MWSchedule -Name $MWName -ApplyTo $ApplyTo -ErrorVariable Error -ErrorAction Stop
        Write-Log -Message "Setting $MWName on $CollectionName - Successful" -SkipEventLog
    }
    Catch {

        #  Write to log in case of failure
        Write-Log -Message "Setting $MWName on $CollectionName - Failed!"
    }
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


#endregion
##*=============================================
##* END FUNCTION LISTINGS
##*=============================================

##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody

## Process Imported CSV Object Data
$csvFileData | ForEach-Object {

    #  Check if we need the Remove Existing Maintenance Window is set
    If ($_.RemoveExisting -eq 'YES' ) {

        #  Write to Log and Console
        $Message = "`n Remove Existing Maintenance Windows from "+$_.CollectionName
        Write-Log -Message $Message -SkipEventLog

        #  Remove Maintenance Window
        Remove-MaintenanceWindows $_.CollectionName
    }

    #  Check if we need to set Maintenance Windows for the whole year
    If ($_.SetForWholeYear -eq "YES") {

        #  Write to Log and Console
        $Message =  "`n `n Setting Maintenance Windows on"+$_.CollectionName+" for the whole year:"
        Write-Log -Message $Message -SkipEventLog

        #  Set Maintenance Windows for the whole year
        For ($Month = [int]$_.Month; $Month -le 12; $Month++) {
            Set-MaintenanceWindows -CollectionName $_.CollectionName -Year $_.Year -Month $Month -OffsetWeeks $_.OffsetWeeks -OffsetDays $_.OffsetDays -StartTime $_.StartTime -StopTime $_.StopTime -ApplyTo $_.ApplyTo
        }

    #  Run without removing Maintenance Windows and set Maintenance Window just for one month
    } Else {

            #  Write to Log and Console
            $Message = "`n `n Setting Maintenance Windows on "+$_.CollectionName+" for "+(Get-Culture).DateTimeFormat.GetMonthName($_.Month)+":"
            Write-Log -Message $Message -SkipEventLog

            #  Set Maintenance Windows
            Set-MaintenanceWindows -CollectionName $_.CollectionName -Year $_.Year -Month $_.Month -OffsetWeeks $_.OffsetWeeks -OffsetDays $_.OffsetDays -StartTime $_.StartTime -StopTime $_.StopTime -ApplyTo $_.ApplyTo
        }
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
#Send-Mail -Body $ResultString

## Write to FileLog
Write-Log -Message "Processing Finished..." -SkipEventLog

## Return to Script Path
Set-Location $ScriptPath

## Remove SCCM PSH Module
Remove-Module "ConfigurationManager" -Force -ErrorAction 'Continue'

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
