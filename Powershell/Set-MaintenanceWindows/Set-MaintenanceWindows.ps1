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

## Cleaning prompt history
CLS

## Get Running Path
$CurrentDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)

## Import Modules and Files
$MWCSVData = Import-Csv $CurrentDir\"Set-MaintenanceWindows.csv"
Import-Module "E:\SCCM\AdminConsole\bin\ConfigurationManager.psd1"

## Set Global variables
$MonthArray = New-Object System.Globalization.DateTimeFormatInfo
$MonthNames = $MonthArray.MonthNames

## Log Initialization
$LogPath = "C:\Temp\MW"
$ErrorLog = $LogPath+"\MWError.log"
If ((Test-Path $LogPath) -eq $False) {
    New-Item -Path $LogPath -Type Directory | Out-Null
} ElseIf (Test-Path $LogPath) {
        Remove-Item $LogPath\* -Recurse -Force
    }
#  Clean Log
Get-Date | Out-File $ErrorLog -Force

## Change Path to CM Site
CD VSM:

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
    Write messages to a log file in CMTrace.exe compatible format or Legacy text file format.
.DESCRIPTION
    Write messages to a log file in CMTrace.exe compatible format or Legacy text file format and optionally display in the console.
.PARAMETER Message
    The message to write to the log file or output to the console.
.EXAMPLE
    Write-Log -Message "Error"
.NOTES
.LINK
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Text')]
        [string]$Message
    )

    ## Getting the Date and time
    $DateAndTime = Get-Date

    ### Writing to log file
    "$DateAndTime : $Message" | Out-File $ErrorLog -Append
    "$DateAndTime : $_" | Out-File $ErrorLog -Append

    ## Writing to Console
    Write-Host $Message  -ForegroundColor Red -BackgroundColor White
    Write-Host $_.Exception -ForegroundColor White -BackgroundColor Red
    Write-Host ""
}
#endregion

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
            Write-Host $_.Name " - Removed!"  -ForegroundColor Green
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
.PARAMETER MWYear
    The Maintenance Window Start Year.
.PARAMETER MWMonth
    The Maintenance Window Start Month.
.PARAMETER MWOffsetWeeks
    The Maintenance Window offset number of weeks after Patch Tuesday.
.PARAMETER MWOffsetWeeks
    The Maintenance Window offset number of days after Tuesday.
.PARAMETER MWStartTime
    The Maintenance Window Start Time.
.PARAMETER MWStopTime
    The Maintenance Window Stop Time.
.EXAMPLE
    Remove-MaintenanceWindows -CollectionName "Computer Collection" -MWYear 2015 -MWMonth 3 -MWOffsetWeeks 3 -MWOffsetDays 2 -MWStartTime "01:00:00"  -MWStopTime "02:00:00"
.NOTES
.LINK
#>
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Collection')]
        [string]$CollectionName,
        [Parameter(Mandatory=$true,Position=1)]
        [Alias('Year')]
        [int16]$MWYear,
        [Parameter(Mandatory=$true,Position=2)]
        [Alias('Month')]
        [int16]$MWMonth,
        [Parameter(Mandatory=$true,Position=3)]
        [Alias('Weeks')]
        [int16]$MWOffsetWeeks,
        [Parameter(Mandatory=$true,Position=4)]
        [Alias('Days')]
        [int16]$MWOffsetDays,
        [Parameter(Mandatory=$true,Position=5)]
        [Alias('StartTime')]
        [string]$MWStartTime,
        [Parameter(Mandatory=$true,Position=6)]
        [Alias('StopTime')]
        [string]$MWStopTime
    )

    ## Get CollectionID
    Try {
        $CollectionID = (Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop -ErrorVariable Error).CollectionID
    }
    Catch {

        #  Write to log in case of failure
        Write-Log "Getting $CollectionName ID - Failed!"
    }

    ## Get PatchTuesday
    $PatchTuesday = Get-PatchTuesday $MWYear $MWMonth

    ## Setting Patch Day, adding offset days and weeks. Get-Date is used to get the date in the same cast, otherwise we cannot convert the date from string to datetime format.
    $PatchDay = (Get-Date -Date $PatchTuesday).AddDays($MWOffsetDays+($MWOffsetWeeks*7))

    ## Check if we got ourselves in the next year and return to the main script if true
    If ($PatchDay.Year -gt $MWYear) {
        Write-Warning "Year threshold detected! Ending Cycle..."
        Return
    }

    ## Setting Maintenance Window Start and Stop times
    $MWStartTime = (Get-Date -Format "M/dd/yyyy HH:mm:ss" -Date $PatchDay) -Replace "00:00:00", $MWStartTime
    $MWStopTime = (Get-Date -Format "M/dd/yyyy HH:mm:ss" -Date $PatchDay) -Replace "00:00:00", $MWStopTime

    ## Create The Schedule Token
    $MWSchedule = New-CMSchedule -Start $MWStartTime -End $MWStopTime -NonRecurring

    ## Set Maintenance Window Naming Convention MW Month and Actual Patches being deployed Month
    $MWName =  'MW.NR.'+(Get-Date -Uformat %B $MWStartTime)+'.Patch_Group_'+$MWOffsetWeeks+'.'+$MonthNames[$MWMonth-1]+'_Updates'

    ## Set Maintenance Window
    Try {
        $SetNewMW = New-CMMaintenanceWindow -CollectionID $CollectionID -Schedule $MWSchedule -Name $MWName -ApplyTo Any -ErrorVariable Error -ErrorAction Stop
        Write-Host "Setting $MWName on $CollectionName - Successful" -ForegroundColor Green
    }
    Catch {

        #  Write to log in case of failure
        Write-Log "Setting $MWName on $CollectionName - Failed!"
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
$MWCSVData | ForEach-Object {

    #  Check if we need the Remove Existing Maintenance Window is set
    If ($_.RemoveExistingMW -eq 1) {
        Write-Host ""
        Write-Host "Remove Existing Maintenance Windows from" $_.CollectionName":" -ForegroundColor Blue -BackgroundColor White
        Write-Host ""
        Remove-MaintenanceWindows $_.CollectionName
    }

    #  Check if we need to set Maintenance Windows for the whole year
    If ($_.SetForWholeYear -eq 1) {
        Write-Host ""
        Write-Host "Setting Maintenance Windows on" $_.CollectionName "for the whole year:" -ForegroundColor Blue -BackgroundColor White
        Write-Host ""
        For ($Month = [int]$_.MWMonth; $Month -le 12; $Month++) {
            Set-MaintenanceWindows -CollectionName $_.CollectionName -MWYear $_.MWYear -MWMonth $Month -MWOffsetWeeks $_.MWOffsetWeeks -MWOffsetDays $_.MWOffsetDays -MWStartTime $_.MWStartTime -MWStopTime $_.MWStopTime
        }

    #  Run without removing Maintenance Windows and set Maintenance Window just for one month
    } Else {
            Write-Host ""
            Write-Host "Setting Maintenance Windows on" $_.CollectionName "for " $MonthNames[$_.MWMonth-1]":" -ForegroundColor Blue -BackgroundColor White
            Write-Host ""
            Set-MaintenanceWindows -CollectionName $_.CollectionName -MWYear $_.MWYear -MWMonth $_.MWMonth -MWOffsetWeeks $_.MWOffsetWeeks -MWOffsetDays $_.MWOffsetDays -MWStartTime $_.MWStartTime -MWStopTime $_.MWStopTime
        }
}
Write-Host ""
Write-Host "Processing Finished!" -BackgroundColor Green -ForegroundColor White

## Return to Script Path
CD $CurrentDir

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
