<#
************************************************************************************************************
* Requires          | Requires PowerShell 3.0                                                              *
* ======================================================================================================== *
* Modified by       |    Date    | Revision | Comments                                                     *
* ________________________________________________________________________________________________________ *
* Ioan Popovici     | 2015-11-13 | v1.0     | First version                                                *
* Ioan Popovici     | 2015-11-16 | v1.1     | Improved logging                                             *
* Ioan Popovici     | 2015-11-17 | v1.2     | Vastly improved                                              *
* Ioan Popovici     | 2016-02-03 | v2.0     | Vastly improved                                              *
* Ioan Popovici     | 2016-02-04 | v2.1     | Fixed TotalSize decimals                                     *
* Ioan Popovici     | 2016-02-19 | v2.2     | EventLog logging support                                     *
* Ioan Popovici     | 2016-02-20 | v2.3     | Added check for not downloaded Cache Items, improved logging *
* Ioan Popovici     | 2017-04-26 | v2.4     | Basic error management, formatting cleanup                   *
* Ioan Popovici     | 2017-04-26 | v2.5     | Orphaned cache cleanup, null ContentID fix, improved logging *
* Ioan Popovici     | 2017-05-02 | v2.5     | Basic error Management                                       *
* Walker            | 2017-08-08 | v2.6     | Fixed first time run logging bug                             *
* Ioan Popovici     | 2018-07-05 | v3.0     | Completely re-written, and optimized se notes                *
* Ioan Popovici     | 2018-07-09 | v3.0     | Added ReferencedThreshold, squashed lots of bugs             *
* Ioan Popovici     | 2018-07-10 | v3.1     | Fixed should run bug                                         *
* Ioan Popovici     | 2018-08-07 | v3.2     | Fixed division by 0 and added basic debug info               *
* ======================================================================================================== *
*                                                                                                          *
************************************************************************************************************

.SYNOPSIS
    Cleans the configuration manager client cache.
.DESCRIPTION
    Cleans the configuration manager client cache of all unneeded with the option to delete persisted content.
.PARAMETER CleanupActions
    Specifies cleanup action to perform. ('All', 'Applications', 'Packages', 'Updates', 'Orphaned'). Default is: 'All'.
    If it's set to 'All' all cleaning actions will be performed.
.PARAMETER LowDiskSpaceThreshold
    Specifies the low disk space threshold percentage after which the cache is cleaned. Default is: '100'.
    If it's set to '100' Free Space Threshold Percentage is ignored.
.PARAMETER ReferencedThreshold
    Specifies to remove cache element only if it has not been referenced in specified number of days. Default is: 0.
    If it's set to '0' Last Referenced Time is ignored.
.PARAMETER LoggingOptions
    Specifies logging options: ('Host', 'File', 'EventLog', 'None'). Default is: ('Host', 'File', 'EventLog').
.PARAMETER SkipSuperPeer
    This switch specifies to skip cleaning if the client is a super peer (Peer Cache). Default is: $false.
.PARAMETER RemovePersisted
    This switch specifies to remove content even if it's persisted. Default is: $false
.EXAMPLE
    PowerShell.exe .\Clean-CMClientCache -CleanupActions "Applications, Packages, Updates, Orphaned" -LoggingOptions 'Host' -LowDiskSpaceThreshold '100' -ReferencedThreshold '30' -SkipSuperPeer -RemovePersisted  -Verbose -Debug
.INPUTS
    None
.OUTPUTS
    System.Management.Automation.PSObject
.NOTES
    ## Improvements
        * Added better logging and logging options by adapting the PADT logging cmdlet. (Slightly modified version)
        * Added support for verbose and debug to the PADT logging cmdlet.
        * Added more cleaning options.
        * Added LowDiskSpaceThreshold option to only clean cache when there is not enough space on the disk.
        * Added SkipSuperPeer, for Peer Cache 'Hosts'.
        * Added ReferencedThreshold, for skipping cache younger than specified number of days.
    ## Fixes
        * Fixed persisted cache cleaning, it's not removed without the RemovePersisted switch.
        * Fixed orphaned cache cleaning and it's not a hack anymore.
        * Fixed error reporting.
    ## Optimizations
        * Speed.
        * The functionality is now split correctly in functions.
        * Script is now ConfigurationItem friendly.
        * Cmdlets are now module friendly.
        * Moved file log in $Env:WinDir\Logs\Clean-CMClientCache.
    ## For issue reporting please use github
        [MyGithub](https://github.com/JhonnyTerminus/SCCMZone/issues)
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

## Get script parameters
Param (
    [Parameter(Mandatory=$false,Position=0)]
    [ValidateSet('All', 'Applications', 'Packages', 'Updates', 'Orphaned')]
    [Alias('Action')]
    [string[]]$CleanupActions = 'All',
    [Parameter(Mandatory=$false,Position=1)]
    [ValidateNotNullorEmpty()]
    [Alias('FreeSpace')]
    [int16]$LowDiskSpaceThreshold = '100',
    [Parameter(Mandatory=$false,Position=2)]
    [ValidateNotNullorEmpty()]
    [Alias('OlderThan')]
    [int16]$ReferencedThreshold = '0',
    [Parameter(Mandatory=$false,Position=3)]
    [ValidateSet('Host', 'File', 'EventLog', 'None')]
    [Alias('Logging')]
    [string[]]$LoggingOptions = @('Host', 'File', 'EventLog'),
    [Parameter(Mandatory=$false,Position=4)]
    [switch]$SkipSuperPeer = $false,
    [Parameter(Mandatory=$false,Position=5)]
    [switch]$RemovePersisted = $false
)

## Initialize result variable
[psobject]$CleanupResult = @()

## Set script variables
$script:LoggingOptions = $LoggingOptions
$script:ReferencedThreshold = $ReferencedThreshold
#  Initialize ShouldRun with true. It will be checked in the script body
[boolean]$ShouldRun = $true

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Resolve-Error
Function Resolve-Error {
<#
.SYNOPSIS
    Enumerate error record details.
.DESCRIPTION
    Enumerate an error record, or a collection of error record, properties. By default, the details for the last error will be enumerated.
.PARAMETER ErrorRecord
    The error record to resolve. The default error record is the latest one: $global:Error[0]. This parameter will also accept an array of error records.
.PARAMETER Property
    The list of properties to display from the error record. Use "*" to display all properties.
    Default list of error properties is: Message, FullyQualifiedErrorId, ScriptStackTrace, PositionMessage, InnerException
.PARAMETER GetErrorRecord
    Get error record details as represented by $_.
.PARAMETER GetErrorInvocation
    Get error record invocation information as represented by $_.InvocationInfo.
.PARAMETER GetErrorException
    Get error record exception details as represented by $_.Exception.
.PARAMETER GetErrorInnerException
    Get error record inner exception details as represented by $_.Exception.InnerException. Will retrieve all inner exceptions if there is more than one.
.EXAMPLE
    Resolve-Error
.EXAMPLE
    Resolve-Error -Property *
.EXAMPLE
    Resolve-Error -Property InnerException
.EXAMPLE
    Resolve-Error -GetErrorInvocation:$false
.NOTES
    Unmodified version of the PADT error resolving cmdlet. I did not write the original cmdlet, please do not credit me for it!
.LINK
    https://psappdeploytoolkit.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [AllowEmptyCollection()]
        [array]$ErrorRecord,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [string[]]$Property = ('Message','InnerException','FullyQualifiedErrorId','ScriptStackTrace','PositionMessage'),
        [Parameter(Mandatory=$false,Position=2)]
        [switch]$GetErrorRecord = $true,
        [Parameter(Mandatory=$false,Position=3)]
        [switch]$GetErrorInvocation = $true,
        [Parameter(Mandatory=$false,Position=4)]
        [switch]$GetErrorException = $true,
        [Parameter(Mandatory=$false,Position=5)]
        [switch]$GetErrorInnerException = $true
    )

    Begin {
        ## If function was called without specifying an error record, then choose the latest error that occurred
        If (-not $ErrorRecord) {
            If ($global:Error.Count -eq 0) {
                #Write-Warning -Message "The `$Error collection is empty"
                Return
            }
            Else {
                [array]$ErrorRecord = $global:Error[0]
            }
        }

        ## Allows selecting and filtering the properties on the error object if they exist
        [scriptblock]$SelectProperty = {
            Param (
                [Parameter(Mandatory=$true)]
                [ValidateNotNullorEmpty()]
                $InputObject,
                [Parameter(Mandatory=$true)]
                [ValidateNotNullorEmpty()]
                [string[]]$Property
            )

            [string[]]$ObjectProperty = $InputObject | Get-Member -MemberType '*Property' | Select-Object -ExpandProperty 'Name'
            ForEach ($Prop in $Property) {
                If ($Prop -eq '*') {
                    [string[]]$PropertySelection = $ObjectProperty
                    Break
                }
                ElseIf ($ObjectProperty -contains $Prop) {
                    [string[]]$PropertySelection += $Prop
                }
            }
            Write-Output -InputObject $PropertySelection
        }

        #  Initialize variables to avoid error if 'Set-StrictMode' is set
        $LogErrorRecordMsg = $null
        $LogErrorInvocationMsg = $null
        $LogErrorExceptionMsg = $null
        $LogErrorMessageTmp = $null
        $LogInnerMessage = $null
    }
    Process {
        If (-not $ErrorRecord) { Return }
        ForEach ($ErrRecord in $ErrorRecord) {
            ## Capture Error Record
            If ($GetErrorRecord) {
                [string[]]$SelectedProperties = & $SelectProperty -InputObject $ErrRecord -Property $Property
                $LogErrorRecordMsg = $ErrRecord | Select-Object -Property $SelectedProperties
            }

            ## Error Invocation Information
            If ($GetErrorInvocation) {
                If ($ErrRecord.InvocationInfo) {
                    [string[]]$SelectedProperties = & $SelectProperty -InputObject $ErrRecord.InvocationInfo -Property $Property
                    $LogErrorInvocationMsg = $ErrRecord.InvocationInfo | Select-Object -Property $SelectedProperties
                }
            }

            ## Capture Error Exception
            If ($GetErrorException) {
                If ($ErrRecord.Exception) {
                    [string[]]$SelectedProperties = & $SelectProperty -InputObject $ErrRecord.Exception -Property $Property
                    $LogErrorExceptionMsg = $ErrRecord.Exception | Select-Object -Property $SelectedProperties
                }
            }

            ## Display properties in the correct order
            If ($Property -eq '*') {
                #  If all properties were chosen for display, then arrange them in the order the error object displays them by default.
                If ($LogErrorRecordMsg) { [array]$LogErrorMessageTmp += $LogErrorRecordMsg }
                If ($LogErrorInvocationMsg) { [array]$LogErrorMessageTmp += $LogErrorInvocationMsg }
                If ($LogErrorExceptionMsg) { [array]$LogErrorMessageTmp += $LogErrorExceptionMsg }
            }
            Else {
                #  Display selected properties in our custom order
                If ($LogErrorExceptionMsg) { [array]$LogErrorMessageTmp += $LogErrorExceptionMsg }
                If ($LogErrorRecordMsg) { [array]$LogErrorMessageTmp += $LogErrorRecordMsg }
                If ($LogErrorInvocationMsg) { [array]$LogErrorMessageTmp += $LogErrorInvocationMsg }
            }

            If ($LogErrorMessageTmp) {
                $LogErrorMessage = 'Error Record:'
                $LogErrorMessage += "`n-------------"
                $LogErrorMsg = $LogErrorMessageTmp | Format-List | Out-String
                $LogErrorMessage += $LogErrorMsg
            }

            ## Capture Error Inner Exception(s)
            If ($GetErrorInnerException) {
                If ($ErrRecord.Exception -and $ErrRecord.Exception.InnerException) {
                    $LogInnerMessage = 'Error Inner Exception(s):'
                    $LogInnerMessage += "`n-------------------------"

                    $ErrorInnerException = $ErrRecord.Exception.InnerException
                    $Count = 0

                    While ($ErrorInnerException) {
                        [string]$InnerExceptionSeperator = '~' * 40

                        [string[]]$SelectedProperties = & $SelectProperty -InputObject $ErrorInnerException -Property $Property
                        $LogErrorInnerExceptionMsg = $ErrorInnerException | Select-Object -Property $SelectedProperties | Format-List | Out-String

                        If ($Count -gt 0) { $LogInnerMessage += $InnerExceptionSeperator }
                        $LogInnerMessage += $LogErrorInnerExceptionMsg

                        $Count++
                        $ErrorInnerException = $ErrorInnerException.InnerException
                    }
                }
            }

            If ($LogErrorMessage) { $Output = $LogErrorMessage }
            If ($LogInnerMessage) { $Output += $LogInnerMessage }

            Write-Output -InputObject $Output

            If (Test-Path -LiteralPath 'variable:Output') { Clear-Variable -Name 'Output' }
            If (Test-Path -LiteralPath 'variable:LogErrorMessage') { Clear-Variable -Name 'LogErrorMessage' }
            If (Test-Path -LiteralPath 'variable:LogInnerMessage') { Clear-Variable -Name 'LogInnerMessage' }
            If (Test-Path -LiteralPath 'variable:LogErrorMessageTmp') { Clear-Variable -Name 'LogErrorMessageTmp' }
        }
    }
    End {
    }
}
#endregion

#region Function Write-Log
Function Write-Log {
<#
.SYNOPSIS
    Write messages to a log file in CMTrace.exe compatible format or Legacy text file format.
.DESCRIPTION
    Write messages to a log file in CMTrace.exe compatible format or Legacy text file format and optionally display in the console.
.PARAMETER Message
    The message to write to the log file or output to the console.
.PARAMETER Severity
    Defines message type. When writing to console or CMTrace.exe log format, it allows highlighting of message type.
    Options: 1 = Information (default), 2 = Warning (highlighted in yellow), 3 = Error (highlighted in red)
.PARAMETER Source
    The source of the message being logged. Also used as the event log source.
.PARAMETER ScriptSection
    The heading for the portion of the script that is being executed. Default is: $script:installPhase.
.PARAMETER LogType
    Choose whether to write a CMTrace.exe compatible log file or a Legacy text log file.
.PARAMETER LoggingOptions
    Choose where to log 'Console', 'File', 'EventLog' or 'None'. You can choose multiple options.
.PARAMETER LogFileDirectory
    Set the directory where the log file will be saved.
.PARAMETER LogFileName
    Set the name of the log file.
.PARAMETER MaxLogFileSizeMB
    Maximum file size limit for log file in megabytes (MB). Default is 10 MB.
.PARAMETER LogName
    Set the name of the event log.
.PARAMETER EventID
    Set the event id for the event log entry.
.PARAMETER WriteHost
    Write the log message to the console.
.PARAMETER ContinueOnError
    Suppress writing log message to console on failure to write message to log file. Default is: $true.
.PARAMETER PassThru
    Return the message that was passed to the function
.PARAMETER VerboseMessage
    Specifies that the message is a debug message. Verbose messages only get logged if -LogDebugMessage is set to $true.
.PARAMETER DebugMessage
    Specifies that the message is a debug message. Debug messages only get logged if -LogDebugMessage is set to $true.
.PARAMETER LogDebugMessage
    Debug messages only get logged if this parameter is set to $true in the config XML file.
.EXAMPLE
    Write-Log -Message "Installing patch MS15-031" -Source 'Add-Patch' -LogType 'CMTrace'
.EXAMPLE
    Write-Log -Message "Script is running on Windows 8" -Source 'Test-ValidOS' -LogType 'Legacy'
.NOTES
    Slightly modified version of the PSADT logging cmdlet. I did not write the original cmdlet, please do not credit me for it.
.LINK
    https://psappdeploytoolkit.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [AllowEmptyCollection()]
        [Alias('Text')]
        [string[]]$Message,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateRange(1,3)]
        [int16]$Severity = 1,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNull()]
        [string]$Source = 'Unknown',
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [string]$ScriptSection = $script:RunPhase,
        [Parameter(Mandatory=$false,Position=4)]
        [ValidateSet('CMTrace','Legacy')]
        [string]$LogType = 'CMTrace',
        [Parameter(Mandatory=$false,Position=5)]
        [ValidateSet('Host', 'File', 'EventLog', 'None')]
        [string[]]$LoggingOptions = $script:LoggingOptions,
        [Parameter(Mandatory=$false,Position=6)]
        [ValidateNotNullorEmpty()]
        [string]$LogFileDirectory = $(Join-Path -Path $Env:WinDir -ChildPath '\Logs\Clean-CMClientCache'),
        [Parameter(Mandatory=$false,Position=7)]
        [ValidateNotNullorEmpty()]
        [string]$LogFileName = 'Clean-CMClientCache.log',
        [Parameter(Mandatory=$false,Position=8)]
        [ValidateNotNullorEmpty()]
        [int]$MaxLogFileSizeMB = '4',
        [Parameter(Mandatory=$false,Position=9)]
        [ValidateNotNullorEmpty()]
        [string]$LogName = 'Configuration Manager',
        [Parameter(Mandatory=$false,Position=10)]
        [ValidateNotNullorEmpty()]
        [int32]$EventID = 1,
        [Parameter(Mandatory=$false,Position=11)]
        [ValidateNotNullorEmpty()]
        [boolean]$ContinueOnError = $false,
        [Parameter(Mandatory=$false,Position=12)]
        [switch]$PassThru = $false,
        [Parameter(Mandatory=$false,Position=13)]
        [switch]$VerboseMessage = $false,
        [Parameter(Mandatory=$false,Position=14)]
        [switch]$DebugMessage = $false,
        [Parameter(Mandatory=$false,Position=15)]
        [boolean]$LogDebugMessage = $false
    )

    Begin {
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        ## Logging Variables
        #  Log file date/time
        [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
        [string]$LogDate = (Get-Date -Format 'dd-MM-yyyy').ToString()
        If (-not (Test-Path -LiteralPath 'variable:LogTimeZoneBias')) { [int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes }
        [string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
        #  Initialize variables
        [boolean]$WriteHost = $false
        [boolean]$WriteFile = $false
        [boolean]$WriteEvent = $false
        [boolean]$DisableLogging = $false
        [boolean]$ExitLoggingFunction = $false
        If (('Host' -in $LoggingOptions) -and (-not ($VerboseMessage -or $DebugMessage))) { $WriteHost = $true }
        If ('File' -in $LoggingOptions) { $WriteFile = $true }
        If ('EventLog' -in $LoggingOptions) { $WriteEvent = $true }
        If ('None' -in $LoggingOptions) { $DisableLogging = $true }
        #  Check if the script section is defined
        [boolean]$ScriptSectionDefined = [boolean](-not [string]::IsNullOrEmpty($ScriptSection))
        #  Get the file name of the source script
        Try {
            If ($script:MyInvocation.Value.ScriptName) {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
            }
            Else {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
            }
        }
        Catch {
            $ScriptSource = ''
        }
        #  Check if the event log and event source exit
        [boolean]$LogNameNotExists = (-not [System.Diagnostics.EventLog]::Exists($LogName))
        [boolean]$LogSourceNotExists = (-not [System.Diagnostics.EventLog]::SourceExists($Source))

        ## Create script block for generating CMTrace.exe compatible log entry
        [scriptblock]$CMTraceLogString = {
            Param (
                [string]$lMessage,
                [string]$lSource,
                [int16]$lSeverity
            )
            "<![LOG[$lMessage]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$lSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$lSeverity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"
        }

        ## Create script block for writing log entry to the console
        [scriptblock]$WriteLogLineToHost = {
            Param (
                [string]$lTextLogLine,
                [int16]$lSeverity
            )
            If ($WriteHost) {
                #  Only output using color options if running in a host which supports colors.
                If ($Host.UI.RawUI.ForegroundColor) {
                    Switch ($lSeverity) {
                        3 { Write-Host -Object $lTextLogLine -ForegroundColor 'Red' -BackgroundColor 'Black' }
                        2 { Write-Host -Object $lTextLogLine -ForegroundColor 'Yellow' -BackgroundColor 'Black' }
                        1 { Write-Host -Object $lTextLogLine }
                    }
                }
                #  If executing "powershell.exe -File <filename>.ps1 > log.txt", then all the Write-Host calls are converted to Write-Output calls so that they are included in the text log.
                Else {
                    Write-Output -InputObject $lTextLogLine
                }
            }
        }

        ## Create script block for writing log entry to the console as verbose or debug message
        [scriptblock]$WriteLogLineToHostAdvanced = {
            Param (
                [string]$lTextLogLine
            )
            #  Only output using color options if running in a host which supports colors.
            If ($Host.UI.RawUI.ForegroundColor) {
                If ($VerboseMessage) {
                    Write-Verbose -Message $lTextLogLine
                }
                Else {
                    Write-Debug -Message $lTextLogLine
                }
            }
            #  If executing "powershell.exe -File <filename>.ps1 > log.txt", then all the Write-Host calls are converted to Write-Output calls so that they are included in the text log.
            Else {
                Write-Output -InputObject $lTextLogLine
            }
        }

        ## Create script block for event writing log entry
        [scriptblock]$WriteToEventLog = {
            If ($WriteEvent) {
                $EventType = Switch ($Severity) {
                    3 { 'Error' }
                    2 { 'Warning' }
                    1 { 'Information' }
                }
                If ($LogNameNotExists -or $LogSourceNotExists) {
                    Try {
                        #  Create event log
                        $null = New-EventLog -LogName $LogName -Source $Source -ErrorAction 'Stop'
                    }
                    Catch {
                        [boolean]$ExitLoggingFunction = $true
                        #  If error creating directory, write message to console
                        If (-not $ContinueOnError) {
                            Write-Host -Object "[$LogDate $LogTime] [${CmdletName}] $ScriptSection :: Failed to create the event log [$LogName`:$Source]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                        }
                    }
                }
                Try {
                    #  Write to event log
                    Write-EventLog -LogName $LogName -Source $Source -EventId $EventID -EntryType $EventType -Message $ConsoleLogLine -ErrorAction 'Stop'
                }
                Catch {
                    [boolean]$ExitLoggingFunction = $true
                    #  If error creating directory, write message to console
                    If (-not $ContinueOnError) {
                        Write-Host -Object "[$LogDate $LogTime] [${CmdletName}] $ScriptSection :: Failed to write to event log [$LogName`:$Source]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                    }
                }
            }
        }

        ## Exit function if it is a debug message and logging debug messages is not enabled in the config XML file
        If (($DebugMessage -or $VerboseMessage) -and (-not $LogDebugMessage)) { [boolean]$ExitLoggingFunction = $true; Return }
        ## Exit function if logging to file is disabled and logging to console host is disabled
        If (($DisableLogging) -and (-not $WriteHost)) { [boolean]$ExitLoggingFunction = $true; Return }
        ## Exit Begin block if logging is disabled
        If ($DisableLogging) { Return }

        ## Create the directory where the log file will be saved
        If (-not (Test-Path -LiteralPath $LogFileDirectory -PathType 'Container')) {
            Try {
                $null = New-Item -Path $LogFileDirectory -Type 'Directory' -Force -ErrorAction 'Stop'
            }
            Catch {
                [boolean]$ExitLoggingFunction = $true
                #  If error creating directory, write message to console
                If (-not $ContinueOnError) {
                    Write-Host -Object "[$LogDate $LogTime] [${CmdletName}] $ScriptSection :: Failed to create the log directory [$LogFileDirectory]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                }
                Return
            }
        }

        ## Assemble the fully qualified path to the log file
        [string]$LogFilePath = Join-Path -Path $LogFileDirectory -ChildPath $LogFileName
    }
    Process {

        ForEach ($Msg in $Message) {
            ## If the message is not $null or empty, create the log entry for the different logging methods
            [string]$CMTraceMsg = ''
            [string]$ConsoleLogLine = ''
            [string]$LegacyTextLogLine = ''
            If ($Msg) {
                #  Create the CMTrace log message
                If ($ScriptSectionDefined) { [string]$CMTraceMsg = "[$ScriptSection] :: $Msg" }

                #  Create a Console and Legacy "text" log entry
                [string]$LegacyMsg = "[$LogDate $LogTime]"
                If ($ScriptSectionDefined) { [string]$LegacyMsg += " [$ScriptSection]" }
                If ($Source) {
                    [string]$ConsoleLogLine = "$LegacyMsg [$Source] :: $Msg"
                    Switch ($Severity) {
                        3 { [string]$LegacyTextLogLine = "$LegacyMsg [$Source] [Error] :: $Msg" }
                        2 { [string]$LegacyTextLogLine = "$LegacyMsg [$Source] [Warning] :: $Msg" }
                        1 { [string]$LegacyTextLogLine = "$LegacyMsg [$Source] [Info] :: $Msg" }
                    }
                }
                Else {
                    [string]$ConsoleLogLine = "$LegacyMsg :: $Msg"
                    Switch ($Severity) {
                        3 { [string]$LegacyTextLogLine = "$LegacyMsg [Error] :: $Msg" }
                        2 { [string]$LegacyTextLogLine = "$LegacyMsg [Warning] :: $Msg" }
                        1 { [string]$LegacyTextLogLine = "$LegacyMsg [Info] :: $Msg" }
                    }
                }
            }

            ## Execute script block to write the log entry to the console as verbose or debug message
            & $WriteLogLineToHostAdvanced -lTextLogLine $ConsoleLogLine -lSeverity $Severity

            ## Exit function if logging is disabled
            If ($ExitLoggingFunction) { Return }

            ## Execute script block to create the CMTrace.exe compatible log entry
            [string]$CMTraceLogLine = & $CMTraceLogString -lMessage $CMTraceMsg -lSource $Source -lSeverity $lSeverity

            ## Choose which log type to write to file
            If ($LogType -ieq 'CMTrace') {
                [string]$LogLine = $CMTraceLogLine
            }
            Else {
                [string]$LogLine = $LegacyTextLogLine
            }

            ## Write the log entry to the log file and event log if logging is not currently disabled
            If (-not $DisableLogging -and $WriteFile) {
                ## Write to file log
                Try {
                    $LogLine | Out-File -FilePath $LogFilePath -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Stop'
                }
                Catch {
                    If (-not $ContinueOnError) {
                        Write-Host -Object "[$LogDate $LogTime] [$ScriptSection] [${CmdletName}] :: Failed to write message [$Msg] to the log file [$LogFilePath]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                    }
                }
                ## Write to event log
                Try {
                    & $WriteToEventLog -lMessage $ConsoleLogLine -lName $LogName -lSource $Source -lSeverity $Severity
                }
                Catch {
                    If (-not $ContinueOnError) {
                        Write-Host -Object "[$LogDate $LogTime] [$ScriptSection] [${CmdletName}] :: Failed to write message [$Msg] to the log file [$LogFilePath]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                    }
                }
            }

            ## Execute script block to write the log entry to the console if $WriteHost is $true and $LogLogDebugMessage is not $true
            & $WriteLogLineToHost -lTextLogLine $ConsoleLogLine -lSeverity $Severity
        }
    }
    End {
        ## Archive log file if size is greater than $MaxLogFileSizeMB and $MaxLogFileSizeMB > 0
        Try {
            If ((-not $ExitLoggingFunction) -and (-not $DisableLogging)) {
                [IO.FileInfo]$LogFile = Get-ChildItem -LiteralPath $LogFilePath -ErrorAction 'Stop'
                [decimal]$LogFileSizeMB = $LogFile.Length/1MB
                If (($LogFileSizeMB -gt $MaxLogFileSizeMB) -and ($MaxLogFileSizeMB -gt 0)) {
                    ## Change the file extension to "lo_"
                    [string]$ArchivedOutLogFile = [IO.Path]::ChangeExtension($LogFilePath, 'lo_')
                    [hashtable]$ArchiveLogParams = @{ ScriptSection = $ScriptSection; Source = ${CmdletName}; Severity = 2; LogFileDirectory = $LogFileDirectory; LogFileName = $LogFileName; LogType = $LogType; MaxLogFileSizeMB = 0; WriteHost = $WriteHost; ContinueOnError = $ContinueOnError; PassThru = $false }

                    ## Log message about archiving the log file
                    $ArchiveLogMessage = "Maximum log file size [$MaxLogFileSizeMB MB] reached. Rename log file to [$ArchivedOutLogFile]."
                    Write-Log -Message $ArchiveLogMessage @ArchiveLogParams

                    ## Archive existing log file from <filename>.log to <filename>.lo_. Overwrites any existing <filename>.lo_ file. This is the same method SCCM uses for log files.
                    Move-Item -LiteralPath $LogFilePath -Destination $ArchivedOutLogFile -Force -ErrorAction 'Stop'

                    ## Start new log file and Log message about archiving the old log file
                    $NewLogMessage = "Previous log file was renamed to [$ArchivedOutLogFile] because maximum log file size of [$MaxLogFileSizeMB MB] was reached."
                    Write-Log -Message $NewLogMessage @ArchiveLogParams
                }
            }
        }
        Catch {
            ## If renaming of file fails, script will continue writing to log file even if size goes over the max file size
        }
        Finally {
            If ($PassThru) { Write-Output -InputObject $Message }
        }
    }
}
#endregion

#region Function Get-CCMCachedApplications
Function Get-CCMCachedApplications {
<#
.SYNOPSIS
    Lists all ccm cached applications.
.DESCRIPTION
    Lists all configuration manager client cached applications with custom properties.
.EXAMPLE
    Get-CCMCachedApplications
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param ()
    Begin {
        Try {

            ## Set script phase for logging
            $script:RunPhase = 'Processing'

            ## Get the name of this function and write verbose header
            [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            #  Write verbose header
            Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

            ## Initialize the CCM resource manager com object
            [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'

            ## Get ccm cache info
            $CacheInfo = $($CCMComObject.GetCacheInfo().GetCacheElements())

            ## Get ccm application list
            $Applications = Get-CimInstance -Namespace 'Root\ccm\ClientSDK' -ClassName 'CCM_Application' -Verbose:$false

            ## Count the applications
            $ApplicationCount = $($Applications | Measure-Object).Count

            ## Initialize counter
            $ProgressCounter = 0

            ## Initialize result object
            [psobject]$CachedApps = @()
        }
        Catch {
            Write-Log -Message "Initialization failed. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Initialization failed. `n$($_.Exception.Message)"
        }
    }
    Process {
        Try {

            ## Get cached application info
            ForEach ($Application in $Applications) {

                ## Show the progress
                $ProgressCounter++
                Write-Progress -Activity 'Processing Applications' -CurrentOperation $Application.FullName -PercentComplete (($ProgressCounter / $ApplicationCount) * 100)

                ## Get application deployment types
                $ApplicationDTs = ($Application | Get-CimInstance -Verbose:$false).AppDTs

                ## Get application content ID
                ForEach ($DeploymentType in $ApplicationDTs) {
                    #  Assemble Invoke-Method arguments
                    $Arguments = [hashtable]@{
                        'AppDeliveryTypeID' = [string]$($DeploymentType.ID)
                        'Revision' = [UINT32]$($DeploymentType.Revision)
                        'ActionType' = 'Install'
                    }
                    #  Get app content ID via GetContentInfo wmi method
                    $AppContentID = (Invoke-CimMethod -Namespace 'Root\ccm\cimodels' -ClassName 'CCM_AppDeliveryType' -MethodName 'GetContentInfo' -Arguments $Arguments -Verbose:$false).ContentID

                    ## Get the cache info for the application using the ContentID
                    $AppCacheInfo = $CacheInfo | Where-Object { $_.ContentID -eq  $AppContentID }

                    ## If the application is in the cache, assemble properties and add it to the result object
                    If (($AppCacheInfo.ContentSize) -gt 0) {
                        #  Set content size to 0 if null to avoid division by 0
                        If (-not $AppCacheInfo.ContentSize) { $ContentSize = 0 }
                        #  Assemble result object props
                        $CachedAppProps = [ordered]@{
                            Name = $($Application.Name)
                            DeploymentType = $($DeploymentType.Name)
                            InstallState = $($Application.InstallState)
                            ContentID = $($AppCacheInfo.ContentID)
                            ContentVersion = $($AppCacheInfo.ContentVersion)
                            ReferenceCount = $($AppCacheInfo.ReferenceCount)
                            LastReferenceTime = $($AppCacheInfo.LastReferenceTime)
                            Location = $($AppCacheInfo.Location)
                            'Size(MB)' = '{0:N2}' -f $($AppCacheInfo.ContentSize / 1KB)
                            CacheElementID = $($AppCacheInfo.CacheElementID)
                        }
                        #  Add items to result object
                        $CachedApps += New-Object 'PSObject' -Property $CachedAppProps
                    }
                }
            }
        }
        Catch {
            Write-Log -Message "Could not get cached application [$($Application.Name)].  `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not get cached application [$($Application.Name)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $CachedApps
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
    }
}
#endregion

#region Function Get-CCMCachedPackages
Function Get-CCMCachedPackages {
<#
.SYNOPSIS
    Lists all ccm cached packages.
.DESCRIPTION
    Lists all configuration manager client cached packages with custom properties.
.EXAMPLE
    Get-CCMCachedPackages
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param ()
    Begin {
        Try {

            ## Set script phase for logging
            $script:RunPhase = 'Processing'

            ## Get the name of this function and write verbose header
		    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            #  Write verbose header
            Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

            ## Initialize the CCM resource manager com object
            [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'

            ## Get ccm cache info
            $CacheInfo = $($CCMComObject.GetCacheInfo().GetCacheElements())

            ## Get ccm package list
            $Packages = Get-CimInstance -Namespace 'Root\ccm\ClientSDK' -ClassName 'CCM_Program' -Verbose:$false

            ## Count the packages
            $PackageCount = $($Packages | Measure-Object).Count

            ## Initialize counter
            $ProgressCounter = 0

            ## Initialize result object
            [psobject]$CachedPkgs = @()
        }
        Catch {
            Write-Log -Message "Initialization failed. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Throw "Initialization failed. `n$($_.Exception.Message)"
        }
    }
    Process {
        Try {

            ## Get cached package info
            ForEach ($Package in $Packages) {

                ## Show the progress
                $ProgressCounter++
                Write-Progress -Activity 'Processing Packages' -CurrentOperation $($Package.FullName) -PercentComplete $(($ProgressCounter / $PackageCount) * 100)

                ## Debug info
                Write-Log -Message "PowerShell version: $($PSVersionTable.PSVersion | Out-String)" -DebugMessage -Source ${CmdletName}
                Write-Log -Message "CacheInfo: `n $($CacheInfo | Out-String)" -DebugMessage -Source ${CmdletName}
                Write-Log -Message "CurentPackage: `n $($Package | Out-String)" -DebugMessage -Source ${CmdletName}
                Write-Log -Message "Size: `n $($PkgCacheInfo.ContentSize | Out-String)" -DebugMessage -Source ${CmdletName}

                ## Get the cache info for the package using the ContentID
                $PkgCacheInfo = $CacheInfo | Where-Object { $_.ContentID -eq  $Package.PackageID }

                ## Debug info
                Write-Log -Message "CachedInfo: `n $($PkgCacheInfo | Out-String)" -DebugMessage -Source ${CmdletName}

                ## If the package is in the cache, assemble properties and add it to the result object
                If (($PkgCacheInfo.ContentSize) -gt 0) {
                    #  Assemble result object props
                    $CachedPkgProps = [ordered]@{
                        Name = $($Package.FullName)
                        Program = $($Package.Name)
                        LastRunStatus = $($Package.LastRunStatus)
                        RepeatRunBehavior = $($Package.RepeatRunBehavior)
                        ContentID = $($PkgCacheInfo.ContentID)
                        ContentVersion = $($PkgCacheInfo.ContentVersion)
                        ReferenceCount = $($PkgCacheInfo.ReferenceCount)
                        LastReferenceTime = $($PkgCacheInfo.LastReferenceTime)
                        Location = $($PkgCacheInfo.Location)
                        'Size(MB)' = '{0:N2}' -f $($PkgCacheInfo.ContentSize / 1KB)
                        CacheElementID = $($PkgCacheInfo.CacheElementID)
                    }
                    #  Add items to result object
                    $CachedPkgs += New-Object 'PSObject' -Property $CachedPkgProps
                }
            }
        }
        Catch {
            Write-Log -Message "Could not get cached package [$($Package.Name)]. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not get cached package [$($Package.Name)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $CachedPkgs
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
    }
}
#endregion

#region Function Get-CCMCachedUpdates
Function Get-CCMCachedUpdates {
<#
.SYNOPSIS
    Lists all ccm cached updates.
.DESCRIPTION
    Lists all configuration manager client cached updates with custom properties.
.EXAMPLE
    Get-CCMCachedUpdates
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

    [CmdletBinding()]
    Param ()
    Begin {
        Try {

            ## Set script phase for logging
            $script:RunPhase = 'Processing'

            ## Get the name of this function and write verbose header
		    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            #  Write verbose header
            Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

            ## Initialize the CCM resource manager com object
            [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'

            ## Get ccm cache info
            $CacheInfo = $($CCMComObject.GetCacheInfo().GetCacheElements())

            ## Get ccm update list
            $Updates = Get-CimInstance -Namespace 'Root\ccm\SoftwareUpdates\UpdatesStore' -ClassName 'CCM_UpdateStatus' -Verbose:$false

            ## Count the updates
            $UpdateCount = $($Updates | Measure-Object).Count

            ## Initialize counter
            $ProgressCounter = 0

            ## Initialize result object
            [psobject]$CachedUpdates = @()
        }
        Catch {
            Write-Log -Message "Initialization failed. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Throw "Initialization failed. `n$($_.Exception.Message)"
        }
    }
    Process {
        Try {

            ## Get cached update info
            ForEach ($Update in $Updates) {

                ## Show the progress
                $ProgressCounter++
                Write-Progress -Activity 'Processing Updates' -CurrentOperation $Update.FullName -PercentComplete (($ProgressCounter / $UpdateCount) * 100)

                ## Get the cache info for the update using the ContentID
                $UpdateCacheInfo = $CacheInfo | Where-Object { $_.ContentID -eq  $Update.UniqueID }

                ## If the update is in the cache, assemble properties and add it to the result object
                If (($UpdateCacheInfo.ContentSize) -gt 0) {
                    #  Assemble result object props
                    $CachedUpdateProps = [ordered]@{
                        Name = $($Update.Title)
                        Article = $($Update.Article)
                        Status = $($Update.Status)
                        ContentID = $($UpdateCacheInfo.ContentID)
                        ContentVersion = $($UpdateCacheInfo.ContentVersion)
                        ReferenceCount = $($UpdateCacheInfo.ReferenceCount)
                        LastReferenceTime = $($UpdateCacheInfo.LastReferenceTime)
                        Location = $($UpdateCacheInfo.Location)
                        'Size(MB)' = '{0:N2}' -f $($UpdateCacheInfo.ContentSize / 1KB)
                        CacheElementID = $($UpdateCacheInfo.CacheElementID)
                    }
                    #  Add items to result object
                    $CachedUpdates += New-Object 'PSObject' -Property $CachedUpdateProps
                }
            }
        }
        Catch {
            Write-Log -Message "Could not get cached update [$($Update.Title)]. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not get cached update [$($Update.Title)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $CachedUpdates
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
    }
}
#endregion

#region Function Remove-CCMCacheElement
Function Remove-CCMCacheElement {
<#
.SYNOPSIS
    Removes a ccm cache element.
.DESCRIPTION
    Removes a configuration manager client cache element and optionally removes persisted content.
.PARAMETER ContentID
    Specifies the cache content ID to be deleted.
.PARAMETER RemovePersisted
    Specifies to remove cache element even if it's persisted. Default is: $false.
.PARAMETER ReferencedThreshold
    Specifies to remove cache element only if it has not been referenced in the last specified number of days.
    Default is: $script:ReferencedThreshold.
.EXAMPLE
    Remove-CCMCacheElement -ContentID '234234234' -RemovePersisted
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [Alias('ID')]
        [string]$ContentID,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [Alias('RPer')]
        [boolean]$RemovePersisted = $false,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [Alias('DaysThreshold')]
        [int16]$ReferencedThreshold = $script:ReferencedThreshold
    )

    Begin {

        ## Set script phase for logging
        $script:RunPhase = 'Removal'

        ## Get the name of this function and write verbose header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        #  Write verbose header
        Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

        ## Initialize the CCM resource manager com object
        [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'

        ## Initialize result object
        [psobject]$RemovedCache = @()

        ## Set the date threshold
        [datetime]$OlderThan = (Get-Date).AddDays(-$ReferencedThreshold)
    }
    Process {
        Try {

            ## Get the CacheElementIDs to delete
            $CacheInfo = $CCMComObject.GetCacheInfo().GetCacheElements() | Where-Object { ($_.ContentID -eq  $ContentID) }

            ## Write verbose message
            Write-Log -Message "ContentID [$ContentID]" -VerboseMessage -Source ${CmdletName}

            ## Delete cache items (This loop is probably not needed but since I don't know if there can be multiple cache items with the same ContentID...)
            ForEach ($CacheItem in $CacheInfo) {
                #  Delete only if no action is in progress
                If ($CacheItem.ReferenceCount -lt 1) {
                    #  Delete only if the $ReferencedThreshold is respected
                    If ([datetime]($CacheItem.LastReferenceTime) -le $OlderThan) {
                        #  Call the remove cache item method
                        $null = $CCMComObject.GetCacheInfo().DeleteCacheElementEx([string]$($CacheItem.CacheElementID), [bool]$RemovePersisted)
                    }
                    Else {
                        $AboveReferencedThreshold = $true
                    }

                    ## Check if the CacheElement has been deleted
                    $CacheInfo = $CCMComObject.GetCacheInfo().GetCacheElements() | Where-Object { $_.ContentID -eq  $ContentID }
                    #  If cache item still exists perform additional checks (this is a hack it would be nice to get the deployment flags from somewhere)
                    If ($CacheInfo) {
                        #  If cache is above referenced threshold set status to 'AboveReferencedThreshold'
                        If ($AboveReferencedThreshold)  { $RemovalStatus = 'AboveReferencedThreshold' }
                        #  If the RemovePersisted switch is set throw error
                        ElseIf ($RemovePersisted) { Throw "Failed to remove cache element [$($CacheItem.ContentID)]" }
                        #  If cache item still exists and RemovePersisted is not specified set the RemovalStatus to 'Persisted'
                        Else { $RemovalStatus = 'Persisted' }
                    }
                    #  If the cache is no longer present set the status to 'Removed'
                    Else { $RemovalStatus = 'Removed' }
                }
                Else {
                    ## If the cache item is still referenced set the removal status to 'Referenced'
                    $RemovalStatus = 'Referenced'
                }

                ## Build result object
                $RemovedCacheProps = [ordered]@{
                    ContentID = $($CacheItem.ContentID)
                    ContentVersion = $($CacheItem.ContentVersion)
                    ReferenceCount = $($CacheItem.ReferenceCount)
                    LastReferenceTime = $($CacheItem.LastReferenceTime)
                    Location = $($CacheItem.Location)
                    ContentSize = '{0:N2}' -f $($CacheItem.ContentSize /1KB)
                    CacheElementID = $($CacheItem.CacheElementID)
                    RemovalStatus = $RemovalStatus
                }

                ##  Add items to result object
                $RemovedCache += New-Object 'PSObject' -Property $RemovedCacheProps
            }
        }
        Catch {
            Write-Log -Message "Could not delete cache element [$($CacheItem.CacheElementID)]. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not delete cache element [$($CacheItem.CacheElementID)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $RemovedCache
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
    }
}
#endregion

#region Function Remove-CCMCachedApplications
Function Remove-CCMCachedApplications {
<#
.SYNOPSIS
    Removes all ccm cached applications.
.DESCRIPTION
    Removes all ccm cached applications with the option to skip persisted content.
.PARAMETER RemovePersisted
    Specifies to remove cached application even if it's persisted. Default is: $false.
.EXAMPLE
    Remove-CCMCachedApplications -RemovePersisted $true
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [Alias('RPer')]
        [boolean]$RemovePersisted = $false
    )

    Begin {
        Try {

            ## Set script phase for logging
            $script:RunPhase = 'Removal'

            ## Get the name of this function and write verbose header
            [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            #  Write verbose header
            Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

            ## Get ccm cached applications
            $CachedApplications = Get-CCMCachedApplications

            ## Initialize result object
            [psobject]$RemovedApplications = @()
        }
        Catch {
            Write-Log -Message "Initialization failed. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Initialization failed. `n$($_.Exception.Message)"
        }
    }
    Process {
        Try {

            ## Process remove cached applications
            ForEach ($Application in $CachedApplications) {
                #  Call Remove-CCMCacheElement
                $RemoveCacheElement = Remove-CCMCacheElement -ContentID $($Application.ContentID) -RemovePersisted $RemovePersisted

                ## Process deleted cache results (This loop is probably not needed but since I don't know if there can be multiple cache items with the same ContentID...)
                ForEach ($CacheElement in $RemoveCacheElement) {
                    #  Assemble result object props
                    $RemovedApplicationProps = [ordered]@{
                        FullName = $($Application.Name)
                        Name = $($Application.DeploymentType)
                        ContentID = $($CacheElement.ContentID)
                        ContentVersion = $($CacheElement.ContentVersion)
                        ReferenceCount = $($CacheElement.ReferenceCount)
                        LastReferenceTime = $($CacheElement.LastReferenceTime)
                        Location = $($CacheElement.Location)
                        'Size(MB)' = $($CacheElement.ContentSize)
                        CacheElementID = $($CacheElement.CacheElementID)
                        Status = $($CacheElement.RemovalStatus)
                    }
                    #  Add items to result object
                    $RemovedApplications += New-Object 'PSObject' -Property $RemovedApplicationProps
                }
            }
        }
        Catch {
            Write-Log -Message "Could not remove cached application [$($Application.Name)]. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not remove cached application [$($Application.Name)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $RemovedApplications
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
    }
}
#endregion

#region Function Remove-CCMCachedPackages
Function Remove-CCMCachedPackages {
<#
.SYNOPSIS
    Removes all ccm cached packages.
.DESCRIPTION
    Removes all ccm cached packages with the option to skip persisted content.
.PARAMETER RemovePersisted
    Specifies to remove cached package even if it's persisted. Default is: $false.
.EXAMPLE
    Remove-CCMCachedPackages -RemovePersisted $true
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [Alias('RPer')]
        [boolean]$RemovePersisted = $false
    )

    Begin {
        Try {

            ## Set script phase for logging
            $script:RunPhase = 'Removal'

            ## Get the name of this function and write verbose header
		    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            #  Write verbose header
            Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

            ## Get ccm cached packages
            $CachedPackages = Get-CCMCachedPackages

            ## Initialize result object
            [psobject]$RemovePackages = @()
        }
        Catch {
            Write-Log -Message "Initialization failed. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Initialization failed. `n$($_.Exception.Message)"
        }
    }
    Process {
        Try {

            ## Process remove cached packages
            ForEach ($Package in $CachedPackages) {
                #  Check if program in the package needs the cached package, it looked weird to call it Program instead of Package
                If ($Package.LastRunStatus -eq 'Succeeded' -and $Package.RepeatRunBehavior -ne 'RerunAlways' -and $Package.RepeatRunBehavior -ne 'RerunIfSuccess') {
                    #  Call Remove-CCMCacheElement
                    $RemoveCacheElement = Remove-CCMCacheElement -ContentID $($Package.ContentID) -RemovePersisted $RemovePersisted
                }
                Else {
                    $Status = 'Needed'
                }

                ## Process deleted cache results (This loop is probably not needed but since I don't know if there can be multiple cache items with the same ContentID...)
                ForEach ($CacheElement in $RemoveCacheElement) {
                    #  Set removal status
                    If ($Status -ne 'Needed') { $Status = $($CacheElement.RemovalStatus) }
                    #  Assemble result object props
                    $RemovePackageProps = [ordered]@{
                        FullName = $($Package.Name)
                        Name = $($Package.Program)
                        ContentID = $($CacheElement.ContentID)
                        ContentVersion = $($CacheElement.ContentVersion)
                        ReferenceCount = $($CacheElement.ReferenceCount)
                        LastReferenceTime = $($CacheElement.LastReferenceTime)
                        Location = $($CacheElement.Location)
                        'Size(MB)' = $($CacheElement.ContentSize)
                        CacheElementID = $($CacheElement.CacheElementID)
                        Status = $Status
                    }
                    #  Add items to result object
                    $RemovePackages += New-Object 'PSObject' -Property $RemovePackageProps
                }
            }
        }
        Catch {
            Write-Log -Message "Could not remove cached package [$($Package.Name)]. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not remove cached package [$($Package.Name)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $RemovePackages
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
    }
}
#endregion

#region Function Remove-CCMCachedUpdates
Function Remove-CCMCachedUpdates {
<#
.SYNOPSIS
    Removes all ccm cached updates.
.DESCRIPTION
    Removes all ccm cached updates.
.EXAMPLE
    Remove-CCMCachedUpdates
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

    [CmdletBinding()]
    Param ()
    Begin {
        Try {

            ## Set script phase for logging
            $script:RunPhase = 'Removal'

            ## Get the name of this function and write verbose header
            [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            #  Write verbose header
            Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

            ## Get ccm cached updates
            $CachedUpdates = Get-CCMCachedUpdates

            ## Initialize result object
            [psobject]$RemoveUpdates = @()
        }
        Catch {
            Write-Log -Message "Initialization failed. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Throw "Initialization failed. `n$($_.Exception.Message)"
        }
    }
    Process {
        Try {

            ## Process remove cached updates
            ForEach ($Update in $CachedUpdates) {
                #  Check if update is installed
                If ($Update.Status -eq 'Installed') {
                    #  Call Remove-CCMCacheElement
                    $RemoveCacheElement = Remove-CCMCacheElement -ContentID $($Update.ContentID) -RemovePersisted $RemovePersisted
                }
                Else {
                    $Status = 'Needed'
                }

                ## Process deleted cache results (This loop is probably not needed but since I don't know if there can be multiple cache items with the same ContentID...)
                ForEach ($CacheElement in $RemoveCacheElement) {
                    #  Set removal status
                    If ($Status -ne 'Needed') { $Status = $($CacheElement.RemovalStatus) }
                    #  Assemble result object props
                    $RemoveUpdateProps = [ordered]@{
                        FullName = $($Update.Title)
                        Name = $($Update.Article)
                        ContentID = $($CacheElement.ContentID)
                        ContentVersion = $($CacheElement.ContentVersion)
                        ReferenceCount = $($CacheElement.ReferenceCount)
                        LastReferenceTime = $($CacheElement.LastReferenceTime)
                        Location = $($CacheElement.Location)
                        'Size(MB)' = $($CacheElement.ContentSize)
                        CacheElementID = $($CacheElement.CacheElementID)
                        Status = $Status
                    }
                    #  Add items to result object
                    $RemoveUpdates += New-Object 'PSObject' -Property $RemoveUpdateProps
                }
            }
        }
        Catch {
            Write-Log -Message "Could not remove cached update [$($Update.Title)]. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not remove cached update [$($Update.Title)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $RemoveUpdates
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
    }
}
#endregion

#region Function Remove-CCMOrphanedCache
Function Remove-CCMOrphanedCache {
<#
.SYNOPSIS
    Removes all orphaned ccm cache items.
.DESCRIPTION
    Removes all ccm cache items not present it wmi.
.EXAMPLE
    Remove-CCMOrphanedCache
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

    [CmdletBinding()]
    Param ()
    Begin {
        Try {

            ## Set script phase for logging
            $script:RunPhase = 'Removal'

            ## Get the name of this function and write verbose header
            [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            #  Write verbose header
            Write-Log -Message 'Start' -VerboseMessage -Source ${CmdletName}

            ## Initialize the CCM resource manager com object
            [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'

            ## Get ccm disk cache info
            [string]$DiskCachePath = $($CCMComObject.GetCacheInfo()).Location
            $DiskCacheInfo = Get-ChildItem -LiteralPath $DiskCachePath | Select-Object -Property 'FullName', 'Name'

            ## Get ccm cache info
            $WmiCachePaths = $($CCMComObject.GetCacheInfo().GetCacheElements()).Location

            ## Initialize result object
            [psobject]$RemoveOrphaned = @()
        }
        Catch {
            Write-Log -Message "Initialization failed. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Initialization failed. `n$($_.Exception.Message)"
        }
    }
    Process {
        Try {

            ## Process cache items
            ForEach ($CacheElement in $DiskCacheInfo) {
                ## Set variables
                #  Set cache Path
                $CacheElementPath = $($CacheElement.FullName)
                #  Set cache Size
                $CacheElementSize = $(Get-ChildItem -LiteralPath $CacheElementPath -Recurse | Measure-Object -Property 'Length' -Sum).Sum

                ## If disk cache path is not present in wmi, delete it
                If ($CacheElementPath -notin $WmiCachePaths) {
                    #  Remove cache item
                    $RemoveCacheElement = Remove-Item -LiteralPath $CacheElementPath -Recurse -Force

                    #  Assemble result object props
                    $RemoveOrphanedProps = [ordered]@{
                        #FullName = 'N/A'
                        #Name = 'N/A'
                        #ContentID = 'N/A'
                        #ContentVersion = 'N/A'
                        #ReferenceCount = 'N/A'
                        #LastReferenceTime = 'N/A'
                        Location = $CacheElementPath
                        'Size(MB)' = '{0:N2}' -f $($CacheElementSize / 1KB)
                        #CacheElementID = 'NA'
                        Status = 'Removed'
                    }
                    #  Add items to result object
                    $RemoveOrphaned += New-Object 'PSObject' -Property $RemoveOrphanedProps
                }
            }
        }
        Catch {
            Write-Log -Message "Could not remove cached item [$($CacheElementPath)]. `n$(Resolve-Error)" -Severity '3' -Source ${CmdletName}
            Throw "Could not remove cached item [$($CacheElementPath)]. `n$($_.Exception.Message)"
        }
        Finally {
            Write-Output -InputObject $RemoveOrphaned
        }
    }
    End {

        ## Write verbose footer
        Write-Log -Message 'Stop' -VerboseMessage -Source ${CmdletName}
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

Try {

    ## Set script phase for logging
    $script:RunPhase = 'Initialization'

    ## Get the file name of the source script
    Try {
        If ($script:MyInvocation.Value.ScriptName) {
            [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
        }
        Else {
            [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
        }
    }
    Catch {
        $ScriptSource = ''
    }

    ## Write Start verbose message
    Write-Log -Message 'Start' -VerboseMessage -Source $ScriptSource

    ## Initialize the CCM resource manager com object
    [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'

    ## Get cache drive free space percentage
    #  Get ccm cache drive location
    [string]$CacheDrive = $($CCMComObject.GetCacheInfo()).Location | Split-Path -Qualifier
    #  Get cache drive info
    $CacheDriveInfo = Get-CimInstance -ClassName 'Win32_LogicalDisk' -Filter "DeviceID='$CacheDrive'" -Verbose:$false
    #  Get cache drive size in GB
    [int16]$DriveSize = $($CacheDriveInfo.Size) / 1GB
    #  Get cache drive free space in GB
    [int16]$DriveFreeSpace = $($CacheDriveInfo.FreeSpace) / 1GB
    #  Calculate percentage
    [int16]$DriveFreeSpacePercentage = ($DriveFreeSpace * 100 / $DriveSize)

    ## Get super peer status
    [boolean]$CanBeSuperPeer = Get-CimInstance -Namespace 'root\ccm\Policy\Machine\ActualConfig' -ClassName 'CCM_SuperPeerClientConfig' -Verbose:$false | Select-Object -ExpandProperty 'CanBeSuperPeer'

    ## Set run condition. If disk free space is above the specified threshold or CanBeSuperPeer is true and SkipSuperPeer is not specified, the script will not run.
    If (($DriveFreeSpacePercentage -gt $LowDiskSpaceThreshold) -or ($CanBeSuperPeer -eq $true -and $SkipSuperPeer)) { $ShouldRun = $false }

    ## Check run condition and stop execution if $ShouldRun is not $true
    If ($ShouldRun) {
        Write-Log -Message 'Should Run test passed' -VerboseMessage -Source $ScriptSource
    }
    Else {
        Write-Log -Message 'Should Run test failed.' -Severity '3' -Source $ScriptSource
        Write-Log -Message "FreeSpace/Threshold [$DriveFreeSpacePercentage`/$LowDiskSpaceThreshold] | IsSuperPeer/SkipSuperPeer [$CanBeSuperPeer`/$SkipSuperPeer]" -DebugMessage -Source $ScriptSource
        Write-Log -Message 'Stop' -VerboseMessage -Source $ScriptSource

        ## Stop execution
        Exit
    }
    Write-Log -Message 'Stop' -Source $ScriptSource -VerboseMessage
}
Catch {
    Write-Log -Message "Script initialization failed. `n$(Resolve-Error)" -Severity '3' -Source $ScriptSource
    Throw "Script initialization failed. $($_.Exception.Message)"
}
Try {

    ## Set script phase for logging
    $script:RunPhase = 'Cleanup'

    ## Write debug action
    Write-Log -Message  "Cleanup Actions [$CleanupActions]" -DebugMessage -Source $ScriptSource

    ## Process selected actions
    Switch ($CleanupActions) {
        All {
            $CleanupResult += Remove-CCMCachedApplications -RemovePersisted $RemovePersisted
            $CleanupResult += Remove-CCMCachedPackages -RemovePersisted $RemovePersisted
            $CleanupResult += Remove-CCMCachedUpdates
            $CleanupResult += Remove-CCMOrphanedCache
        }
        Applications {
            $CleanupResult += Remove-CCMCachedApplications -RemovePersisted $RemovePersisted
        }
        Packages {
            $CleanupResult += Remove-CCMCachedPackages -RemovePersisted $RemovePersisted
        }
        Updates {
            $CleanupResult += Remove-CCMCachedUpdates
        }
        Orphaned {
            $CleanupResult += Remove-CCMOrphanedCache
        }
        Default {
            Write-Log -Message "Invalid cleanup action [$_] selected." -Severity '3' -Source $ScriptSource
            Throw "Invalid cleanup action selected."
        }
    }
}
Catch {
    Write-Log -Message "Could not perform cleanup action. `n$(Resolve-Error)" -Severity '3' -Source $ScriptSource
    Throw "Could not perform cleanup action. `n$($_.Exception.Message)"
}
Finally {

    ## Set script phase for logging
    $script:RunPhase = 'CleanupResult'

    ## Calculate total deleted size
    $TotalDeletedSize = $CleanupResult | Where-Object { $_.Status -eq 'Removed' } | Measure-Object -Property 'Size(MB)' -Sum | Select-Object -ExpandProperty Sum
    If (-not $TotalDeletedSize) { $TotalDeletedSize = 0 }

    ## Assemble output result
    $OutputResult = $($CleanupResult | Format-List -Property FullName, Name, Location, LastReferenceTime, 'Size(MB)', Status | Out-String) + "TotalDeletedSize: " + $TotalDeletedSize

    ## Write output to log, event log and console and status
    Write-Log -Message $OutputResult -Source $ScriptSource

    ## Write verbose stop
    Write-Log -Message 'Stop' -VerboseMessage -Source $ScriptSource
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================