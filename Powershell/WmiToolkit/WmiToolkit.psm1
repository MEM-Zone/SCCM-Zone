<#
*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 4.0                                                  *
* ===================================================================================================== *
* Modified by     |    Date    | Revision | Comments                                                    *
* _____________________________________________________________________________________________________ *
* Octavian Cordos | 2017-11-09 | v0.0.1     | First version                                             *
* Ioan Popovici   | 2017-11-09 | v0.0.2     | Created functions                                         *
* Ioan Popovici   | 2017-11-27 | v0.0.3     | All planned functions added                               *
* Ioan Popovici   | 2017-11-27 | v0.0.4     | Fully functional and tested, copy and rename do not work  *
* Ioan Popovici   | 2017-12-17 | v0.0.5     | Complete re-write, to convoluted and overdesigned         *
* Ioan Popovici   | 2017-12-20 | v0.0.6     | Scrapped about 60% should be readable now                 *
* Ioan Popovici   | 2017-12-22 | v0.0.7     | Implemented default PS cmdlet ErrorHandling               *
* Ioan Popovici   | 2017-12-23 | v0.0.8     | Fixed/Simplified input/output where possible              *
* Ioan Popovici   | 2017-12-27 | v0.0.9     | Get/Set/New except Set-WmiInstance, working and tested    *
* Ioan Popovici   | 2017-12-27 | v0.1.0     | Remove-WmiInstance re-written and working                 *
* Ioan Popovici   | 2017-01-07 | v0.1.1     | Get-WmiClass output fix, Remove class, namespace working  *
* Ioan Popovici   | 2017-01-08 | v0.1.2     | Remove class qualifiers working working                   *
* Ioan Popovici   | 2017-01-09 | v0.1.3     | Fixed namespace functions input, only path input now      *
* Ioan Popovici   | 2017-01-09 | v0.1.4     | Fixed Get-WmiPropertyQualifier output                     *
* Ioan Popovici   | 2017-01-09 | v0.1.5     | All Remove functions working correctly now                *
* Ioan Popovici   | 2017-01-10 | v0.1.6     | Fix namespace recurse deletion and creation               *
* Ioan Popovici   | 2017-01-10 | v0.1.7     | Fix Copy-WmiClassQualifiers                               *
* Ioan Popovici   | 2017-01-10 | v0.1.8     | Fix New-WmiClass namespace detection bug                  *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This PowerShell module contains functions for managing WMI.
.DESCRIPTION
    This PowerShell module contains functions for creating WMI Namespaces, Classes and Instances.
.EXAMPLE
    Import-Module -Name 'C:\Temp\WmiToolkit.psm1' -Verbose
.EXAMPLE
    Get-Command -Module 'WmiToolkit'
.NOTES
    --
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings


#region Function Write-FunctionHeaderOrFooter
Function Write-FunctionHeaderOrFooter {
<#
.SYNOPSIS
    Write the function header or footer to the log upon first entering or exiting a function.
.DESCRIPTION
    Write the "Function Start" message, the bound parameters the function was invoked with, or the "Function End" message when entering or exiting a function.
    Messages are debug messages so will only be logged if LogDebugMessage option is enabled in XML config file.
.PARAMETER CmdletName
    The name of the function this function is invoked from.
.PARAMETER CmdletBoundParameters
    The bound parameters of the function this function is invoked from.
.PARAMETER Header
    Write the function header.
.PARAMETER Footer
    Write the function footer.
.EXAMPLE
    Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
.EXAMPLE
    Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://psappdeploytoolkit.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]$CmdletName,
        [Parameter(Mandatory=$true,ParameterSetName='Header')]
        [AllowEmptyCollection()]
        [hashtable]$CmdletBoundParameters,
        [Parameter(Mandatory=$true,ParameterSetName='Header')]
        [switch]$Header,
        [Parameter(Mandatory=$true,ParameterSetName='Footer')]
        [switch]$Footer
    )

    If ($Header) {
        Write-Log -Message 'Function Start' -Source ${CmdletName} -DebugMessage

        ## Get the parameters that the calling function was invoked with
        [string]$CmdletBoundParameters = $CmdletBoundParameters | Format-Table -Property @{ Label = 'Parameter'; Expression = { "[-$($_.Key)]" } }, @{ Label = 'Value'; Expression = { $_.Value }; Alignment = 'Left' } -AutoSize -Wrap | Out-String
        If ($CmdletBoundParameters) {
            Write-Log -Message "Function invoked with bound parameter(s): `n$CmdletBoundParameters" -Source ${CmdletName} -DebugMessage
        }
        Else {
            Write-Log -Message 'Function invoked without any bound parameters.' -Source ${CmdletName} -DebugMessage
        }
    }
    ElseIf ($Footer) {
        Write-Log -Message 'Function End' -Source ${CmdletName} -DebugMessage
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
    The source of the message being logged.
.PARAMETER ScriptSection
    The heading for the portion of the script that is being executed. Default is: $script:installPhase.
.PARAMETER LogType
    Choose whether to write a CMTrace.exe compatible log file or a Legacy text log file.
.PARAMETER LogFileDirectory
    Set the directory where the log file will be saved.
    Default is %WINDIR%\Logs\WmiToolkit.
.PARAMETER LogFileName
    Set the name of the log file.
.PARAMETER MaxLogFileSizeMB
    Maximum file size limit for log file in megabytes (MB). Default is 10 MB.
.PARAMETER WriteHost
    Write the log message to the console.
.PARAMETER ContinueOnError
    Suppress writing log message to console on failure to write message to log file. Default is: $true.
.PARAMETER PassThru
    Return the message that was passed to the function
.PARAMETER DebugMessage
    Specifies that the message is a debug message. Debug messages only get logged if -LogDebugMessage is set to $true.
.PARAMETER LogDebugMessage
    Debug messages only get logged if this parameter is set to $true in the config XML file.
.EXAMPLE
    Write-Log -Message "Installing patch MS15-031" -Source 'Add-Patch' -LogType 'CMTrace'
.EXAMPLE
    Write-Log -Message "Script is running on Windows 8" -Source 'Test-ValidOS' -LogType 'Legacy'
.NOTES
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
        [string]$Source = '',
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [string]$ScriptSection = '',
        [Parameter(Mandatory=$false,Position=4)]
        [ValidateSet('CMTrace','Legacy')]
        [string]$LogType = 'Legacy',
        [Parameter(Mandatory=$false,Position=5)]
        [ValidateNotNullorEmpty()]
        [string]$LogFileDirectory = $(Join-Path -Path $Env:windir -ChildPath "\Logs\WmiToolkit"),
        [Parameter(Mandatory=$false,Position=6)]
        [ValidateNotNullorEmpty()]
        [string]$LogFileName = 'WmiTool.log',
        [Parameter(Mandatory=$false,Position=7)]
        [ValidateNotNullorEmpty()]
        [decimal]$MaxLogFileSizeMB = '5',
        [Parameter(Mandatory=$false,Position=8)]
        [ValidateNotNullorEmpty()]
        [boolean]$WriteHost = $true,
        [Parameter(Mandatory=$false,Position=9)]
        [ValidateNotNullorEmpty()]
        [boolean]$ContinueOnError = $true,
        [Parameter(Mandatory=$false,Position=10)]
        [switch]$PassThru = $false,
        [Parameter(Mandatory=$false,Position=11)]
        [switch]$DebugMessage = $false,
        [Parameter(Mandatory=$false,Position=12)]
        [boolean]$LogDebugMessage = $true
    )

    Begin {
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        ## Logging Variables
        #  Log file date/time
        [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
        [string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
        If (-not (Test-Path -LiteralPath 'variable:LogTimeZoneBias')) { [int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes }
        [string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
        #  Initialize variables
        [boolean]$ExitLoggingFunction = $false
        If (-not (Test-Path -LiteralPath 'variable:DisableLogging')) { $DisableLogging = $false }
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

        ## Exit function if it is a debug message and logging debug messages is not enabled in the config XML file
        If (($DebugMessage) -and (-not $LogDebugMessage)) { [boolean]$ExitLoggingFunction = $true; Return }
        ## Exit function if logging to file is disabled and logging to console host is disabled
        If (($DisableLogging) -and (-not $WriteHost)) { [boolean]$ExitLoggingFunction = $true; Return }
        ## Exit Begin block if logging is disabled
        If ($DisableLogging) { Return }
        ## Exit function function if it is an [Initialization] message and the toolkit has been relaunched
    If ($ScriptSection -eq 'Initialization') { [boolean]$ExitLoggingFunction = $true; Return }

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
        ## Exit function if logging is disabled
        If ($ExitLoggingFunction) { Return }

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

            ## Execute script block to create the CMTrace.exe compatible log entry
            [string]$CMTraceLogLine = & $CMTraceLogString -lMessage $CMTraceMsg -lSource $Source -lSeverity $Severity

            ## Choose which log type to write to file
            If ($LogType -ieq 'CMTrace') {
                [string]$LogLine = $CMTraceLogLine
            }
            Else {
                [string]$LogLine = $LegacyTextLogLine
            }

            ## Write the log entry to the log file if logging is not currently disabled
            If (-not $DisableLogging) {
                Try {
                    $LogLine | Out-File -FilePath $LogFilePath -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Stop'
                }
                Catch {
                    If (-not $ContinueOnError) {
                        Write-Host -Object "[$LogDate $LogTime] [$ScriptSection] [${CmdletName}] :: Failed to write message [$Msg] to the log file [$LogFilePath]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                    }
                }
            }

            ## Execute script block to write the log entry to the console if $WriteHost is $true
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
        [switch]$GetErrorRecord,
        [Parameter(Mandatory=$false,Position=3)]
        [switch]$GetErrorInvocation,
        [Parameter(Mandatory=$false,Position=4)]
        [switch]$GetErrorException,
        [Parameter(Mandatory=$false,Position=5)]
        [switch]$GetErrorInnerException
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


#region Function Get-WmiNameSpace
Function Get-WmiNameSpace {
<#
.SYNOPSIS
    This function is used to get a WMI namespace.
.DESCRIPTION
    This function is used to get the details of one or more WMI namespaces.
.PARAMETER Namespace
    Specifies the namespace path. Supports wildcards.
.EXAMPLE
    Get-WmiNameSpace -NameSpace 'ROOT\SCCM'
.EXAMPLE
    Get-WmiNameSpace -NameSpace 'ROOT\*'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## If namespace is 'ROOT' or -List is specified get namespace else get Parent\Leaf namespace
            If ($List -or ($Namespace -eq 'ROOT')) {
                $GetNamespace = Get-CimInstance -Namespace $Namespace -ClassName '__Namespace'
            }
            Else {
                #  Set namespace path and name
                $NamespaceParent = $(Split-Path -Path $Namespace -Parent)
                $NamespaceLeaf = $(Split-Path -Path $Namespace -Leaf)
                #  Get namespace
                $GetNamespace = Get-CimInstance -Namespace $NamespaceParent -ClassName '__Namespace' | Where-Object { $_.Name -like $NamespaceLeaf }
            }

            ## If no namespace is found, write debug message and optionally throw error is -ErrorAction 'Stop' is specified
            If (-not $GetNamespace) {
                $NamespaceNotFoundErr = "Namespace [$Namespace] not found."
                Write-Log -Message $NamespaceNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $NamespaceNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to retrieve wmi namespace [$Namespace]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $GetNamespace
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Get-WmiClass
Function Get-WmiClass {
<#
.SYNOPSIS
    This function is used to get WMI class details.
.DESCRIPTION
    This function is used to get the details of one or more WMI classes.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name to search for. Supports wildcards. Default is: '*'.
.PARAMETER QualifierName
    Specifies the qualifier name to search for.(Optional)
.PARAMETER IncludeSpecialClasses
    Specifies to include System, MSFT and CIM classes. Use this or Get operations only.
.EXAMPLE
    Get-WmiClass -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone'
.EXAMPLE
    Get-WmiClass -Namespace 'ROOT\SCCM' -QualifierName 'Description'
.EXAMPLE
    Get-WmiClass -Namespace 'ROOT\SCCM'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName = '*',
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$QualifierName,
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [switch]$IncludeSpecialClasses
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the namespace exists
            $NamespaceTest = Get-WmiNameSpace -Namespace $Namespace -ErrorAction 'SilentlyContinue'
            If (-not $NamespaceTest) {
                $NamespaceNotFoundErr = "Namespace [$Namespace] not found."
                Write-Log -Message $NamespaceNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $NamespaceNotFoundErr -Category 'ObjectNotFound'
            }

            ## Get all class details 
            If ($QualifierName) {
                $WmiClass = Get-CimClass -Namespace $Namespace -Class $ClassName -QualifierName $QualifierName -ErrorAction 'SilentlyContinue'
            }
            Else {
                $WmiClass = Get-CimClass -Namespace $Namespace -Class $ClassName -ErrorAction 'SilentlyContinue'
            }

            ## Filter class or classes details based on specified parameters
            If ($IncludeSpecialClasses) {
                $GetClass = $WmiClass
            }
            Else {
                $GetClass = $WmiClass | Where-Object { ($_.CimClassName -notmatch '__') -and ($_.CimClassName -notmatch 'CIM_') -and ($_.CimClassName -notmatch 'MSFT_') }
            }

            ## If no class is found, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $GetClass) {
                $ClassNotFoundErr = "No class [$ClassName] found in namespace [$Namespace]."
                Write-Log -Message $ClassNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $ClassNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to retrieve wmi class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $GetClass
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Get-WmiClassQualifier
Function Get-WmiClassQualifier {
<#
.SYNOPSIS
    This function is used to get the qualifiers of a WMI class.
.DESCRIPTION
    This function is used to get one or more qualifiers of a WMI class.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to get the qualifiers.
.PARAMETER QualifierName
    Specifies the qualifier search for. Suports wildcards. Default is: '*'.
.PARAMETER QualifierValue
    Specifies the qualifier search for. Supports wildcards.(Optional)
.EXAMPLE
    Get-WmiClassQualifier -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -QualifierName 'Description' -QualifierValue 'SCCMZone Blog'
.EXAMPLE
    Get-WmiClassQualifier -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -QualifierName 'Description' -QualifierValue 'SCCMZone*'
.EXAMPLE
    Get-WmiClassQualifier -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$QualifierName = '*',
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [string]$QualifierValue
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Get the all class qualifiers
            $WmiClassQualifier = (Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop' | Select-Object *).CimClassQualifiers | Where-Object -Property Name -like $QualifierName

            ## Filter class qualifiers according to specifed parameters
            If ($QualifierValue) {
                $GetClassQualifier = $WmiClassQualifier | Where-Object -Property Value -like $QualifierValue
            }
            Else {
                $GetClassQualifier = $WmiClassQualifier
            }

            ## If no class qualifiers are found, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $GetClassQualifier) {
                $ClassQualifierNotFoundErr = "No qualifier [$QualifierName] found for class [$Namespace`:$ClassName]."
                Write-Log -Message $ClassQualifierNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $ClassQualifierNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to retrieve wmi class [$Namespace`:$ClassName] qualifier [$QualifierName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $GetClassQualifier
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Get-WmiProperty
Function Get-WmiProperty {
<#
.SYNOPSIS
    This function is used to get the properties of a WMI class.
.DESCRIPTION
    This function is used to get one or more properties of a WMI class.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to get the properties.
.PARAMETER PropertyName
    Specifies the propery name to search for. Supports wildcards. Default is: '*'.
.PARAMETER PropertyValue
    Specifies the propery value or values to search for. Supports wildcards.(Optional)
.PARAMETER QualifierName
    Specifies the property qualifier name to match. Supports wildcards.(Optional)
.PARAMETER Property
    Matches property Name, Value and CimType. Can be piped. If this parameter is specified all other search parameters will be ignored.(Optional)
    Supported format:
        [PSCustomobject]@{
            'Name' = 'Website'
            'Value' = $null
            'CimType' = 'String'
        }
.EXAMPLE
    Get-WmiProperty -Namespace 'ROOT' -ClassName 'SCCMZone'
.EXAMPLE
    Get-WmiProperty -Namespace 'ROOT' -ClassName 'SCCMZone' -PropertyName 'WebsiteSite' -QualifierName 'key'
.EXAMPLE
    Get-WmiProperty -Namespace 'ROOT' -ClassName 'SCCMZone' -PropertyName '*Site'
.EXAMPLE
    $Property = [PSCustomobject]@{
        'Name' = 'Website'
        'Value' = $null
        'CimType' = 'String'
    }
    Get-WmiProperty -Namespace 'ROOT' -ClassName 'SCCMZone' -Property $Property
    $Property | Get-WmiProperty -Namespace 'ROOT' -ClassName 'SCCMZone'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyName = '*',
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyValue,
        [Parameter(Mandatory=$false,Position=4)]
        [ValidateNotNullorEmpty()]
        [string]$QualifierName,
        [Parameter(Mandatory=$false,ValueFromPipeline,Position=5)]
        [ValidateNotNullorEmpty()]
        [PSCustomObject]$Property = @()
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if class exists
            $ClassTest = Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'SilentlyContinue'

            ## If no class is found, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $ClassTest) {
                $ClassNotFoundErr = "No class [$ClassName] found in namespace [$Namespace]."
                Write-Log -Message $ClassNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $ClassNotFoundErr -Category 'ObjectNotFound'
            }

            ## Get class properties
            $WmiProperty = (Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'SilentlyContinue' | Select-Object *).CimClassProperties | Where-Object -Property Name -like $PropertyName

            ## Get class property based on specified parameters
            If ($Property) {

                #  Compare all specified properties and return only properties that match Name, Value and CimType.
                $GetProperty = Compare-Object -ReferenceObject $Property -DifferenceObject $WmiProperty -Property Name, Value, CimType -IncludeEqual -ExcludeDifferent -PassThru

            }
            ElseIf ($PropertyValue -and $QualifierName) {
                $GetProperty = $WmiProperty | Where-Object { ($_.Value -like $PropertyValue) -and ($_.Qualifiers.Name -like $QualifierName) }
            }
            ElseIf ($PropertyValue) {
                $GetProperty = $WmiProperty | Where-Object -Property Value -like $PropertyValue
            }
            ElseIf ($QualifierName) {
                $GetProperty = $WmiProperty | Where-Object { $_.Qualifiers.Name -like $QualifierName }
            }
            Else {
                $GetProperty = $WmiProperty
            }

            ## If no matching properties are found, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $GetProperty) {
                $PropertyNotFoundErr = "No property [$PropertyName] found for class [$Namespace`:$ClassName]."
                Write-Log -Message $PropertyNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $PropertyNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to retrieve wmi class [$Namespace`:$ClassName] properties. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $GetProperty
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Get-WmiPropertyQualifier
Function Get-WmiPropertyQualifier {
<#
.SYNOPSIS
    This function is used to get the property qualifiers of a WMI class.
.DESCRIPTION
    This function is used to get one or more property qualifiers of a WMI class.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to get the property qualifiers.
.PARAMETER PropertyName
    Specifies the property name for which to get the property qualifiers. Supports wilcards. Can be piped. Default is: '*'.
.PARAMETER QualifierName
    Specifies the property qualifier name or names to search for.(Optional)
.PARAMETER QualifierValue
    Specifies the property qualifier value or values to search for.(Optional)
.EXAMPLE
    Get-WmiPropertyQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -PropertyName 'SCCMZone Blog'
.EXAMPLE
    'SCCMZone Blog', 'ServerAddress' | Get-WmiPropertyQualifier -Namespace 'ROOT' -ClassName 'SCCMZone'
.EXAMPLE
    Get-WmiPropertyQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -QualifierName 'key','Description'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,ValueFromPipeline,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyName = '*',
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [string[]]$QualifierName,
        [Parameter(Mandatory=$false,Position=4)]
        [ValidateNotNullorEmpty()]
        [string[]]$QualifierValue
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Get all details for the specified property name
            $WmiPropertyQualifier = (Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop').CimClassProperties | Where-Object -Property Name -like $PropertyName | Select-Object -ExpandProperty 'Qualifiers'

            ## Get property qualifiers based on specified parameters
            If ($QualifierName -and $QualifierValue) {
                $GetPropertyQualifier = $WmiPropertyQualifier | Where-Object { ($_.Name -in $QualifierName) -and ($_.Value -in $QualifierValue) }
            }
            ElseIf ($QualifierName) {
                $GetPropertyQualifier = $WmiPropertyQualifier | Where-Object { ($_.Name -in $QualifierName) }
            }
            ElseIf ($QualifierValue) {
                $GetPropertyQualifier = $WmiPropertyQualifier | Where-Object { $_.Value -in $QualifierValue }
            }
            Else {
                $GetPropertyQualifier = $WmiPropertyQualifier
            }

            ## On property qualifiers retrieval failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $GetPropertyQualifier) {
                $PropertyQualifierNotFoundErr = "No property [$PropertyName] qualifier [$QualifierName `= $QualifierValue] found for class [$Namespace`:$ClassName]."
                Write-Log -Message $PropertyQualifierNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $PropertyQualifierNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to retrieve wmi class [$Namespace`:$ClassName] property qualifiers. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $GetPropertyQualifier
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Get-WmiInstance
Function Get-WmiInstance {
<#
.SYNOPSIS
    This function is used get the values of an WMI instance.
.DESCRIPTION
    This function is used find a WMI instance by comparing properties. It will return the the instance where all specified properties match.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to get the instance properties.
.PARAMETER Property
    Specifies the class instance properties and values to find.
.PARAMETER KeyOnly
    Indicates that only objects with key properties populated are returned.
.EXAMPLE
    [hashtable]$Property = @{
        'ServerPort' = '80'
        'ServerIP' = '10.10.10.11'
        'Source' = 'SCCMZone Blog'
    }
    Get-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Property $Property
.EXAMPLE
    Get-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Property @{ 'Source' = 'SCCMZone Blog' } -KeyOnly
.EXAMPLE
    Get-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [string]$Namespace = 'ROOT\cimv2',
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [hashtable]$Property,
        [Parameter(Mandatory=$false,Position=3)]
        [switch]$KeyOnly
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the class exists
            $null = Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop'

            ## Get all instance details or get only details where the key properties are filled in
            If ($KeyOnly) {
                $WmiInstance = Get-CimInstance -Namespace $Namespace -ClassName $ClassName -KeyOnly
            }
            Else {
                $WmiInstance = Get-CimInstance -Namespace $Namespace -ClassName $ClassName
            }

            ## Match instance details based on specified parameters
            If ($WmiInstance) {
                If ($Property) {

                    #  Get Property Names from function input to be used for filtering
                    [string[]]$InputPropertyNames =  $($Property.GetEnumerator().Name)
        
                    #  Convert Property hashtable to PSCustomObject for comparison
                    [PSCustomObject]$InputProperty = [PSCustomObject]$Property
            
                    #  -ErrorAction 'SilentlyContinue' does not seem to work correctly with the Compare-Object commandlet so it needs to be set globaly
                    $OriginalErrorActionPreference = $ErrorActionPreference
                    $ErrorActionPreference = 'SilentlyContinue'

                    #  Check if and instance with the same values exists. Since $InputProperty is a dinamically generated object Compare-Object has no hope of working correctly.
                    #  Luckily Compare-Object as a -Property parameter which allows us to look at specific parameters.
                    $GetInstance = $WmiInstance | ForEach-Object {
                        $MatchInstance = Compare-Object -ReferenceObject $_ -DifferenceObject $InputProperty -Property $InputPropertyNames -IncludeEqual -ExcludeDifferent
                        If ($MatchInstance) {
                            #  Add matched instance to output
                            $_
                        }
                    }

                    #  Setting the ErrorActionPreference back to the previous value
                    $ErrorActionPreference = $OriginalErrorActionPreference
                }
                Else {
                    $GetInstance = $WmiInstance
                }
            }

            #  If no instances (or matching instances) are found, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $GetInstance) {
                $InstanceNotFoundErr = "No matching instances found in class [$Namespace`:$ClassName]."
                Write-Log -Message $InstanceNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $InstanceNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to retrieve wmi instances for class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $GetInstance
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function New-WmiNameSpace
Function New-WmiNameSpace {
<#
.SYNOPSIS
    This function is used to create a new WMI namespace.
.DESCRIPTION
    This function is used to create a new WMI namespace.
.PARAMETER Namespace
    Specifies the namespace to create.
.PARAMETER CreateSubTree
    This swith is used to create the whole namespace sub tree if it does not exist.
.EXAMPLE
    New-WmiNameSpace -Namespace 'ROOT\SCCM'
.EXAMPLE
    New-WmiNameSpace -Namespace 'ROOT\SCCM\SCCMZone\Blog' -CreateSubTree
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [switch]$CreateSubTree = $false       
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the namespace exists
            $WmiNamespace = Get-WmiNameSpace -Namespace $Namespace -ErrorAction 'SilentlyContinue'
            
            ## Create Namespace if it does not exist
            If (-not $WmiNamespace) {
               
                #  Split path into it's components
                $NamespacePaths = $Namespace.Split('\')

                #  Assigning root namespace, just for show, should always be 'ROOT'
                [string]$Path = $NamespacePaths[0]

                #  Initialize NamespacePathsObject
                [PSCustomObject]$NamespacePathsObject = @()

                #  Parsing path components and assemle individual paths
                For ($i = 1; $i -le $($NamespacePaths.Length -1); $i++ ) {
                    $Path += '\' + $NamespacePaths[$i]

                    #  Assembing path props and add them to the NamspacePathsObject
                    $PathProps = [ordered]@{ Name = $(Split-Path -Path $Path) ; Value = $(Split-Path -Path $Path -Leaf) }
                    $NamespacePathsObject += $PathProps
                }            
                
                #  Split path into it's components
                $NamespacePaths = $Namespace.Split('\')

                #  Assigning root namespace, just for show, should always be 'ROOT'
                [string]$Path = $NamespacePaths[0]

                #  Initialize NamespacePathsObject
                [PSCustomObject]$NamespacePathsObject = @()

                #  Parsing path components and assemle individual paths
                For ($i = 1; $i -le $($NamespacePaths.Length -1); $i++ ) {
                    $Path += '\' + $NamespacePaths[$i]

                    #  Assembing path props and add them to the NamspacePathsObject
                    $PathProps = [ordered]@{ 
                        'NamespacePath' = $(Split-Path -Path $Path) 
                        'NamespaceName' = $(Split-Path -Path $Path -Leaf)
                        'NamespaceTest' = [boolean]$(Get-WmiNameSpace -Namespace $Path -ErrorAction 'SilentlyContinue')
                    }
                    $NamespacePathsObject += [PSCustomObject]$PathProps
                }

                #  If the path does not contain missing subnamespaces or the -CreateSubTree switch is specified create namespace or namespaces
                If (($($NamespacePathsObject -match $false).Count -eq 1 ) -or $CreateSubTree) {

                    #  Create each namespace in path one by one
                    $NamespacePathsObject | ForEach-Object {

                        #  Check if we need to create the namespace
                        If (-not $_.NamespaceTest) {
                            #  Create namespace object and assign namespace name
                            $NameSpaceObject = (New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList "\\.\$($_.NameSpacePath)`:__NAMESPACE").CreateInstance()
                            $NameSpaceObject.Name = $_.NamespaceName
            
                            #  Write the namespace object
                            $NewNamespace = $NameSpaceObject.Put()
                            $NameSpaceObject.Dispose()
                        }
                        Else {
                            Write-Log -Message "Namespace [$($_.NamespacePath)`\$($_.NamespaceName)] already exists." -Severity 2 -Source ${CmdletName} -DebugMessage
                        }
                    }

                    #  On namespace creation failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
                    If (-not $NewNamespace) {
                        $CreateNamespaceErr = "Failed to create namespace [$($_.NameSpacePath)`\$($_.NamespaceName)]."
                        Write-Log -Message $CreateNamespaceErr -Severity 3 -Source ${CmdletName} -DebugMessage
                        Write-Error -Message $CreateNamespaceErr -Category 'InvalidResult'
                    }
                }
                ElseIf (($($NamespacePathsObject -match $false).Count -gt 1)) {
                    $SubNamespaceFoundErr = "Child namespace detected in namespace path [$Namespace]. Use the -CreateSubtree switch to create the whole path."
                    Write-Log -Message $SubNamespaceFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                    Write-Error -Message $SubNamespaceFoundErr -Category 'InvalidOperation'
                }
            }
            Else {
                $NamespaceAlreadyExistsErr = "Failed to create namespace. [$Namespace] already exists."
                Write-Log -Message $NamespaceAlreadyExistsErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $NamespaceAlreadyExistsErr -Category 'ResourceExists'
            }
        }
        Catch {
            Write-Log -Message "Failed to create namespace [$Namespace]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
        }
        Finally {
            Write-Output -InputObject $NewNamespace
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function New-WmiClass
Function New-WmiClass {
<#
.SYNOPSIS
    This function is used to create a WMI class.
.DESCRIPTION
    This function is used to create a WMI class with custom properties.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI namespace. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the name for the new class.
.PARAMETER Qualifiers
    Specifies one ore more property qualifiers using qualifier name and value only. You can omit this parameter or enter one or more items in the hashtable.
    You can also specify a string but you must separate the name and value with a new line character (`n). This parameter can also be piped.
    The qualifiers will be added with these default values and flavors:
        Static = $true
        IsAmended = $false
        PropagatesToInstance = $true
        PropagatesToSubClass = $false
        IsOverridable = $true
.PARAMETER CreateDestination
    This switch is used to create destination namespace.
.EXAMPLE
    [hashtable]$Qualifiers = @{
        Key = $true
        Static = $true
        Description = 'SCCMZone Blog'
    }
    New-WmiClass -Namespace 'ROOT' -ClassName 'SCCMZone' -Qualifiers $Qualifiers
.EXAMPLE
    "Key = $true `n Static = $true `n Description = SCCMZone Blog" | New-WmiClass -Namespace 'ROOT' -ClassName 'SCCMZone'
.EXAMPLE
    New-WmiClass -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -CreateDestination
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,ValueFromPipeline,Position=2)]
        [ValidateNotNullorEmpty()]
        [PSCustomObject]$Qualifiers = @("Static = $true"),
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [switch]$CreateDestination = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the class exists
            $ClassTest = Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'SilentlyContinue'

            ## Check if the namespace exists
            $NamespaceTest = Get-WmiNameSpace -Namespace $Namespace -ErrorAction 'SilentlyContinue'

            ## Create destination namespace if specified, otherwise throw error if -ErrorAction 'Stop' is specified
            If ((-not $NamespaceTest) -and $CreateDestination) {
                $null = New-WmiNameSpace $Namespace -CreateSubTree -ErrorAction 'Stop'
            }
            ElseIf (-not $NamespaceTest) {
                $NamespaceNotFoundErr = "Namespace [$Namespace] does not exist. Use the -CreateDestination switch to create namespace."
                Write-Log -Message $NamespaceNotFoundErr -Severity 3 -Source ${CmdletName}
                Write-Error -Message $NamespaceNotFoundErr -Category 'ObjectNotFound'  
            }

            ## Create class if it does not exist
            If (-not $ClassTest) {

                #  Create class object
                [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:__CLASS",[String]::Empty,$null)
                $ClassObject.Name = $ClassName

                #  Write the class and dispose of the class object
                $NewClass = $ClassObject.Put()
                $ClassObject.Dispose()

                #  On class creation failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
                If (-not $NewClass) {

                    #  Error handling and logging
                    $NewClassErr = "Failed to create class [$ClassName] in namespace [$Namespace]."
                    Write-Log -Message $NewClassErr -Severity 3 -Source ${CmdletName} -DebugMessage
                    Write-Error -Message $NewClassErr -Category 'InvalidResult'
                }

                ## If input qualifier is not a hashtable convert string input to hashtable
                If ($Qualifiers -isnot [hashtable]) {
                    $Qualifiers = $Qualifiers | ConvertFrom-StringData
                }

                ## Set property qualifiers one by one if specified, otherwise set default qualifier name, value and flavors
                If ($Qualifiers) {
                    #  Convert to a hashtable format accepted by Set-WmiClassQualifier. Name = QualifierName and Value = QualifierValue are expected.
                    $Qualifiers.Keys | ForEach-Object {
                        [hashtable]$PropertyQualifier = @{ Name = $_; Value = $Qualifiers.Item($_) }
                        #  Set qualifier
                        $null = Set-WmiClassQualifier -Namespace $Namespace -ClassName $ClassName -Qualifier $PropertyQualifier -ErrorAction 'Stop'
                    }
                }
                Else {
                    $null = Set-WmiClassQualifier -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop'
                }
            }
            Else {
                $ClassAlreadyExistsErr = "Failed to create class [$Namespace`:$ClassName]. Class already exists."
                Write-Log -Message $ClassAlreadyExistsErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $ClassAlreadyExistsErr -Category 'ResourceExists'
            }
        }
        Catch {
            Write-Log -Message "Failed to create class [$ClassName] in namespace [$Namespace]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $NewClass
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Set-WmiClassQualifier
Function Set-WmiClassQualifier {
<#
.SYNOPSIS
    This function is used to set qualifiers to a WMI class.
.DESCRIPTION
    This function is used to set qualifiers to a WMI class. Existing qualifiers with the same name will be overwriten
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI namespace. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to add the qualifiers.
.PARAMETER Qualifier
    Specifies the qualifier name, value and flavours as hashtable. You can omit this parameter or enter one or more items in the hashtable.
    You can also specify a string but you must separate the name and value with a new line character (`n). This parameter can also be piped.
    If you omit a hashtable item the default item value will be used. Only item values can be specified (right of the '=' sign).
    Default is:
        [hashtable][ordered]@{
            Name = 'Static'
            Value = $true
            IsAmended = $false
            PropagatesToInstance = $true
            PropagatesToSubClass = $false
            IsOverridable = $true
        }
.EXAMPLE
    Set-WmiClassQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -Qualifier @{ Name = 'Description'; Value = 'SCCMZone Blog' }
.EXAMPLE
    Set-WmiClassQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -Qualifier "Name = Description `n Value = SCCMZone Blog"
.EXAMPLE
    "Name = Description `n Value = SCCMZone Blog" | Set-WmiClassQualifier -Namespace 'ROOT' -ClassName 'SCCMZone'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false,Position=0)]
    [ValidateNotNullorEmpty()]
    [string]$Namespace = 'ROOT\cimv2',
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateNotNullorEmpty()]
    [string]$ClassName,
    [Parameter(Mandatory=$false,ValueFromPipeline,Position=2)]
    [ValidateNotNullorEmpty()]
    [PSCustomObject]$Qualifier = @()
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the class exist
            $null = Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop'

            ## If input qualifier is not a hashtable convert string input to hashtable
            If ($Qualifier -isnot [hashtable]) {
                $Qualifier = $Qualifier | ConvertFrom-StringData
            }

            ## Add the missing qualifier value, name and flavor to the hashtable using splatting
            If (-not $Qualifier.Item('Name')) { $Qualifier.Add('Name', 'Static') }
            If (-not $Qualifier.Item('Value')) { $Qualifier.Add('Value', $true) }
            If (-not $Qualifier.Item('IsAmended')) { $Qualifier.Add('IsAmended', $false) }
            If (-not $Qualifier.Item('PropagatesToInstance')) { $Qualifier.Add('PropagatesToInstance', $true) }
            If (-not $Qualifier.Item('PropagatesToSubClass')) { $Qualifier.Add('PropagatesToSubClass', $false) }
            If (-not $Qualifier.Item('IsOverridable')) { $Qualifier.Add('IsOverridable', $true) }

            ## Create the ManagementClass object
            [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:$ClassName")

            ## Set key qualifier if specified, otherwise set qualifier
            $ClassObject.Qualifiers.Add($Qualifier.Item('Name'), $Qualifier.Item('Value'), $Qualifier.Item('IsAmended'), $Qualifier.Item('PropagatesToInstance'), $Qualifier.Item('PropagatesToSubClass'), $Qualifier.Item('IsOverridable'))
            $SetClassQualifiers = $ClassObject.Put()
            $ClassObject.Dispose()

            ## On class qualifiers creation failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $SetClassQualifiers) {

                #  Error handling and logging
                $SetClassQualifiersErr = "Failed to set qualifier [$Qualifier.Item('Name')] for class [$Namespace`:$ClassName]."
                Write-Log -Message $SetClassQualifiersErr -Severity 3 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $SetClassQualifiersErr -Category 'InvalidResult'
            }
        }
        Catch {
            Write-Log -Message "Failed to set qualifier for class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $SetClassQualifiers
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function New-WmiProperty
Function New-WmiProperty {
<#
.SYNOPSIS
    This function is used to add properties to a WMI class.
.DESCRIPTION
    This function is used to add custom properties to a WMI class.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI namespace. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to add the properties.
.PARAMETER PropertyName
    Specifies the property name.
.PARAMETER PropertyType
    Specifies the property type.
.PARAMETER Qualifiers
    Specifies one ore more property qualifiers using qualifier name and value only. You can omit this parameter or enter one or more items in the hashtable.
    You can also specify a string but you must separate the name and value with a new line character (`n). This parameter can also be piped.
    The qualifiers will be added with these default flavors:
        IsAmended = $false
        PropagatesToInstance = $true
        PropagatesToSubClass = $false
        IsOverridable = $true
.PARAMETER Key
    Specifies if the property is key. Default is: false.(Optional)
.EXAMPLE
    [hashtable]$Qualifiers = @{
        Key = $true
        Static = $true
        Description = 'SCCMZone Blog'
    }
    New-WmiProperty -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -PropertyName 'Website' -PropertyType 'String' -Qualifiers $Qualifiers
.EXAMPLE
    "Key = $true `n Description = SCCMZone Blog" | New-WmiProperty -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -PropertyName 'Website' -PropertyType 'String'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyName,
        [Parameter(Mandatory=$true,Position=3)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyType,
        [Parameter(Mandatory=$false,ValueFromPipeline,Position=4)]
        [ValidateNotNullorEmpty()]
        [PSCustomObject]$Qualifiers = @()
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the class exists
            $null = Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop'

            ## Check if the property exist
            $WmiPropertyTest = Get-WmiProperty -Namespace $Namespace -ClassName $ClassName -PropertyName $PropertyName -ErrorAction 'SilentlyContinue'

            ## Create the property if it does not exist
            If (-not $WmiPropertyTest) {

                #  Set property to array if specified
                If ($PropertyType -match 'Array') {
                    $PropertyType = $PropertyType.Replace('Array','')
                    $PropertyIsArray = $true
                }
                Else {
                    $PropertyIsArray = $false
                }

                #  Create the ManagementClass object
                [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:$ClassName")

                #  Add class property
                $ClassObject.Properties.Add($PropertyName, [System.Management.CimType]$PropertyType, $PropertyIsArray)

                #  Write class object
                $NewProperty = $ClassObject.Put()
                $ClassObject.Dispose()

                ## On property creation failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
                If (-not $NewProperty) {

                    #  Error handling and logging
                    $NewPropertyErr = "Failed create property [$PropertyName] for Class [$Namespace`:$ClassName]."
                    Write-Log -Message $NewPropertyErr -Severity 3 -Source ${CmdletName} -DebugMessage
                    Write-Error -Message $NewPropertyErr -Category 'InvalidResult'
                }

                ## Set property qualifiers one by one if specified
                If ($Qualifiers) {
                    #  Convert to a hashtable format accepted by Set-WmiPropertyQualifier. Name = QualifierName and Value = QualifierValue are expected.
                    $Qualifiers.Keys | ForEach-Object {
                        [hashtable]$PropertyQualifier = @{ Name = $_; Value = $Qualifiers.Item($_) }
                        #  Set qualifier
                        $null = Set-WmiPropertyQualifier -Namespace $Namespace -ClassName $ClassName -PropertyName $PropertyName -Qualifier $PropertyQualifier -ErrorAction 'Stop'
                    }
                }
            }
            Else {
                $PropertyAlreadyExistsErr = "Property [$PropertyName] already present for class [$Namespace`:$ClassName]."
                Write-Log -Message $PropertyAlreadyExistsErr  -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $PropertyAlreadyExistsErr -Category 'ResourceExists'
            }
        }
        Catch {
            Write-Log -Message "Failed to create property for class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $NewProperty
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Set-WmiPropertyQualifier
Function Set-WmiPropertyQualifier {
<#
.SYNOPSIS
    This function is used to set WMI property qualifier value.
.DESCRIPTION
    This function is used to set WMI property qualifier value to an existing WMI property.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI namespace. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to add the properties.
.PARAMETER PropertyName
    Specifies the property name.
.PARAMETER Qualifier
    Specifies the qualifier name, value and flavours as hashtable. You can omit this parameter or enter one or more items in the hashtable.
    You can also specify a string but you must separate the name and value with a new line character (`n). This parameter can also be piped.
    If you omit a hashtable item the default item value will be used. Only item values can be specified (right of the '=' sign).
    Default is:
        [hashtable][ordered]@{
            Name = 'Static'
            Value = $true
            IsAmended = $false
            PropagatesToInstance = $true
            PropagatesToSubClass = $false
            IsOverridable = $true
        }
    Specifies if the property is key. Default is: $false.
.EXAMPLE
    Set-WmiPropertyQualifier -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -Property 'WebSite' -Qualifier @{ Name = 'Description' ; Value = 'SCCMZone Blog' }
.EXAMPLE
    Set-WmiPropertyQualifier -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -Property 'WebSite' -Qualifier "Name = Description `n Value = SCCMZone Blog"
.EXAMPLE
    "Name = Description `n Value = SCCMZone Blog" | Set-WmiPropertyQualifier -Namespace 'ROOT\SCCM' -ClassName 'SCCMZone' -Property 'WebSite'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyName,
        [Parameter(Mandatory=$false,ValueFromPipeline,Position=3)]
        [ValidateNotNullorEmpty()]
        [PSCustomObject]$Qualifier = @()
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the property exists
            $null = Get-WmiProperty -Namespace $Namespace -ClassName $ClassName -PropertyName $PropertyName -ErrorAction 'Stop'

            ## If input qualifier is not a hashtable convert string input to hashtable
            If ($Qualifier -isnot [hashtable]) {
                $Qualifier = $Qualifier | ConvertFrom-StringData
            }

            ## Add the missing qualifier value, name and flavor to the hashtable using splatting
            If (-not $Qualifier.Item('Name')) { $Qualifier.Add('Name', 'Static') }
            If (-not $Qualifier.Item('Value')) { $Qualifier.Add('Value', $true) }
            If (-not $Qualifier.Item('IsAmended')) { $Qualifier.Add('IsAmended', $false) }
            If (-not $Qualifier.Item('PropagatesToInstance')) { $Qualifier.Add('PropagatesToInstance', $true) }
            If (-not $Qualifier.Item('PropagatesToSubClass')) { $Qualifier.Add('PropagatesToSubClass', $false) }
            If (-not $Qualifier.Item('IsOverridable')) { $Qualifier.Add('IsOverridable', $true) }

            ## Create the ManagementClass object
            [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:$ClassName")

            ## Set key qualifier if specified, otherwise set qualifier
            If ('key' -eq $Qualifier.Item('Name')) {
                $ClassObject.Properties[$PropertyName].Qualifiers.Add('Key', $true)
                $SetClassQualifiers = $ClassObject.Put()
                $ClassObject.Dispose()
            }
            Else {
                $ClassObject.Properties[$PropertyName].Qualifiers.Add($Qualifier.Item('Name'), $Qualifier.Item('Value'), $Qualifier.Item('IsAmended'), $Qualifier.Item('PropagatesToInstance'), $Qualifier.Item('PropagatesToSubClass'), $Qualifier.Item('IsOverridable'))
                $SetClassQualifiers = $ClassObject.Put()
                $ClassObject.Dispose()
            }

            ## On property qualifiers creation failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $SetClassQualifiers) {

                #  Error handling and logging
                $SetClassQualifiersErr = "Failed to set qualifier [$Qualifier.Item('Name')] for property [$Namespace`:$ClassName($PropertyName)]."
                Write-Log -Message $SetClassQualifiersErr -Severity 3 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $SetClassQualifiersErr -Category 'InvalidResult'
            }
        }
        Catch {
            Write-Log -Message "Failed to set property qualifier for class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
        }
        Finally {
            Write-Output -InputObject $SetClassQualifiers
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function New-WmiInstance
Function New-WmiInstance {
<#
.SYNOPSIS
    This function is used to create a WMI Instance.
.DESCRIPTION
    This function is used to create a WMI Instance using CIM.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class where to create the new WMI instance.
.PARAMETER Key
    Specifies properties that are used as keys (Optional).
.PARAMETER Property
    Specifies the class instance Properties or Values. You can also specify a string but you must separate the name and value with a new line character (`n).
    This parameter can also be piped.
.EXAMPLE
    [hashtable]$Property = @{
        'ServerPort' = '89'
        'ServerIP' = '11.11.11.11'
        'Source' = 'File1'
        'Date' = $(Get-Date)
    }
    New-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Key 'File1' -Property $Property
.EXAMPLE
    "Server Port = 89 `n ServerIp = 11.11.11.11 `n Source = File `n Date = $(GetDate)" | New-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Property $Property
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string[]]$Key,
        [Parameter(Mandatory=$true,ValueFromPipeline,Position=3)]
        [ValidateNotNullorEmpty()]
        [PSCustomObject]$Property
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if class exists
            $null = Get-WmiClass -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop'

            ## If input qualifier is not a hashtable convert string input to hashtable
            If ($Property -isnot [hashtable]) {
                $Property = $Property | ConvertFrom-StringData
            }

            ## Create instance
            If ($Key) {
                $NewInstance = New-CimInstance -Namespace $Namespace -ClassName $ClassName -Key $Key -Property $Property
            }
            Else {
                $NewInstance = New-CimInstance -Namespace $Namespace -ClassName $ClassName -Property $Property
            }

            ## On instance creation failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            If (-not $NewInstance) {
                Write-Log -Message "Failed to create instance in class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName} -DebugMessage
            }
        }
        Catch {
            Write-Log -Message "Failed to create instance in class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $NewInstance
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Remove-WmiNameSpace
Function Remove-WmiNameSpace {
<#
.SYNOPSIS
    This function is used to delete a WMI namespace.
.DESCRIPTION
    This function is used to delete a WMI namespace by name.
.PARAMETER Namespace
   Specifies the namespace to remove.
.PARAMETER Force
    This switch deletes all existing classes in the specified path. Default is: $false.
.PARAMETER Recurse
    This switch deletes all existing child namespaces in the specified path.
.EXAMPLE
    Remove-WmiNameSpace -Namespace 'ROOT\SCCM' -Force -Recurse
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace,     
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [switch]$Force = $false,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [switch]$Recurse = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Set namespace root
            $NamespaceRoot = Split-Path -Path $Namespace
            ## Set namespace name
            $NamespaceName = Split-Path -Path $Namespace -Leaf

            ## Check if the namespace exists
            $null = Get-WmiNameSpace -Namespace $Namespace -ErrorAction 'Stop'

            ## Check if there are any classes
            $ClassTest = Get-WmiClass -Namespace $Namespace -ErrorAction 'SilentlyContinue'

            ## Check if there are any child namespaces or if the -Recurse switch was specified
            $ChildNamespaceTest = (Get-WmiNameSpace -Namespace $Namespace'\*' -ErrorAction 'SilentlyContinue').Name
            If ((-not $ChildNamespaceTest) -or $Recurse) {

                #   Remove all existing classes and instances if the -Force switch was specified
                If ($Force -and $ClassTest) {
                    Remove-WmiClass -Namespace $Namespace -RemoveAll
                }
                ElseIf ($ClassTest) {
                    $NamespaceHasClassesErr = "Classes [$($ClassTest.Count)] detected in namespace [$Namespace]. Use the -Force switch to remove classes."
                    Write-Log -Message $NamespaceHasClassesErr -Severity 2 -Source ${CmdletName} -DebugMessage
                    Write-Error -Message $NamespaceHasClassesErr -Category 'InvalidOperation'        
                }

                #  Create the Namespace Object
                $NameSpaceObject = (New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList "\\.\$NamespaceRoot`:__NAMESPACE").CreateInstance()
                $NameSpaceObject.Name = $NamespaceName

                #  Remove the Namespace
                $null = $NameSpaceObject.Delete()
                $NameSpaceObject.Dispose()
            }
            ElseIf ($ChildNamespaceTest) {
                $ChildNamespaceDetectedErr = "Child namespace [$ChildNamespaceTest] detected in namespace [$Namespace]. Use the -Recurse switch to remove Child namespaces."
                Write-Log -Message $ChildNamespaceDetectedErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $ChildNamespaceDetectedErr -Category 'InvalidOperation'  
            }
        }
        Catch {
            Write-Log -Message "Failed to remove namespace [$Namespace]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {}
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Remove-WmiClass
Function Remove-WmiClass {
<#
.SYNOPSISl
    This function is used to remove a WMI class.
.DESCRIPTION
    This function is used to remove a WMI class by name.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name to remove. Can be piped.
.PARAMETER RemoveAll
    This switch is used to remove all namespace classes.
.EXAMPLE
    Remove-WmiClass -Namespace 'ROOT' -ClassName 'SCCMZone','SCCMZoneBlog'
.EXAMPLE
    'SCCMZone','SCCMZoneBlog' | Remove-WmiClass -Namespace 'ROOT'
.EXAMPLE
    Remove-WmiClass -Namespace 'ROOT' -RemoveAll
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$false,ValueFromPipeline,Position=1)]
        [ValidateNotNullorEmpty()]
        [string[]]$ClassName,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [switch]$RemoveAll = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Get classes names
            [string[]]$WmiClassNames = (Get-WmiClass -Namespace $Namespace -ErrorAction 'Stop').CimClassName

            ## Add classes to deletion string array depending on selected options
            If ($RemoveAll) {
                $ClassNamesToDelete = $WmiClassNames
            }
            ElseIf ($ClassName) {
                $ClassNamesToDelete = $WmiClassNames | Where-Object { $_ -in $ClassName }
            }
            Else {
                $ClassNameIsNullErr = "ClassName cannot be `$null if -RemoveAll is not specified."
                Write-Log -Message $ClassNameIsNullErr -Severity 3 -Source ${CmdletName}
                Write-Error -Message $ClassNameIsNullErr -Category 'InvalidArgument'
            }

            ## Remove classes
            If ($ClassNamesToDelete) {
                $ClassNamesToDelete | Foreach-Object {

                    #  Create the class object
                    [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:$_")
                    
                    #  Remove class
                    $null = $ClassObject.Delete()
                    $ClassObject.Dispose()
                }
            }
            Else {
                $ClassNotFoundErr = "No matching class [$ClassName] found for namespace [$Namespace]."
                Write-Log -Message $ClassNotFoundErr -Severity 2 -Source ${CmdletName}
                Write-Error -Message $ClassNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to remove class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {}
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Remove-WmiClassQualifier
Function Remove-WmiClassQualifier {
<#
.SYNOPSIS
    This function is used to remove qualifiers from a WMI class.
.DESCRIPTION
    This function is used to remove qualifiers from a WMI class by name.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI namespace. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to remove the qualifiers.
.PARAMETER QualifierName
    Specifies the qualifier name or names to be removed.
.PARAMETER RemoveAll
    This switch will remove all class qualifiers.
.EXAMPLE
    Remove-WmiClassQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -QualifierName 'Description', 'Static'
.EXAMPLE
    Remove-WmiClassQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -RemoveAll
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false,Position=0)]
    [ValidateNotNullorEmpty()]
    [string]$Namespace = 'ROOT\cimv2',
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateNotNullorEmpty()]
    [string]$ClassName,
    [Parameter(Mandatory=$false,ValueFromPipeline,Position=2)]
    [ValidateNotNullorEmpty()]
    [string[]]$QualifierName,
    [Parameter(Mandatory=$false,Position=3)]
    [ValidateNotNullorEmpty()]
    [switch]$RemoveAll = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Get class qualifiers
            $WmiClassQualifier = (Get-WmiClassQualifier -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop').Name

            ## Add qualifier name to deletion array depending on selected options
            If ($RemoveAll) {
                $RemoveClassQualifier = $WmiClassQualifier
            }
            ElseIf ($QualifierName) {
                $RemoveClassQualifier = $WmiClassQualifier | Where-Object { $_ -in $QualifierName }
            }
            Else {
                $QualifierNameIsNullErr = "QualifierName cannot be `$null if -RemoveAll is not specified."
                Write-Log -Message $QualifierNameIsNullErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $QualifierNameIsNullErr -Category 'InvalidArgument'
            }

            ## Remove qualifiers by name
            If ($RemoveClassQualifier) {
                
                #  Create the ManagementClass object
                [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:$ClassName")

                #  Remove class qualifiers one by one
                $QualifierName | ForEach-Object { $ClassObject.Qualifiers.Remove($_) }

            }
            Else {

                #  Error handling
                $PropertyNotFoundErr = "No matching qualifier [$QualifierName] found for class [$Namespace`:$ClassName]."
                Write-Log -Message $PropertyNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $PropertyNotFoundErr -Category 'ObjectNotFound'
            }
          
            ## Write class object
            $null = $ClassObject.Put()
            $ClassObject.Dispose()
        }
        Catch {
            Write-Log -Message "Failed to remove qualifier for class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {}
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Remove-WmiProperty
Function Remove-WmiProperty {
<#
.SYNOPSIS
    This function is used to remove WMI class properties.
.DESCRIPTION
    This function is used to remove WMI class properties by name.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name for which to remove the properties.
.PARAMETER PropertyName
    Specifies the class property name or names to remove.
.PARAMETER RemoveAll
    This switch is used to remove all properties. Default is: $false. If this switch is specified the Property parameter is ignored.
.PARAMETER Force
    This switch is used to remove all instances. The class must be empty in order to be able to delete properties. Default is: $false.
.EXAMPLE
    Remove-WmiProperty -Namespace 'ROOT' -ClassName 'SCCMZone' -Property 'SCCMZone','Blog'
.EXAMPLE
    Remove-WmiProperty -Namespace 'ROOT' -ClassName 'SCCMZone' -RemoveAll -Force
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string[]]$PropertyName,
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [switch]$RemoveAll = $false,
        [Parameter(Mandatory=$false,Position=4)]
        [ValidateNotNullorEmpty()]
        [switch]$Force = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Get class property names
            [string[]]$WmiPropertyNames = (Get-WmiProperty -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop').Name

            ## Get class instances
            $InstanceTest = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction 'SilentlyContinue'

            ## Add property to deletion string array depending on selected options
            If ($RemoveAll) {
                $RemoveWmiProperty = $WmiPropertyNames
            }
            ElseIf ($PropertyName) {
                $RemoveWmiProperty = $WmiPropertyNames | Where-Object {$_ -in $PropertyName }
            }
            Else {
                $PropertyNameIsNullErr = "PropertyName cannot be `$null if -RemoveAll is not specified."
                Write-Log -Message $PropertyNameIsNullErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $PropertyNameIsNullErr -Category 'InvalidArgument'
            }

            ## Remove class property
            If ($RemoveWmiProperty) {

                #  Remove all existing instances if the -Force switch was specified
                If ($Force -and $InstanceTest) {
                    Remove-WmiInstance -Namespace $Namespace -ClassName $ClassName -RemoveAll -ErrorAction 'Continue'
                }
                ElseIf ($InstanceTest) {
                    $ClassHasInstancesErr = "Instances [$($InstanceTest.Count)] detected in class [$Namespace`:$ClassName]. Use the -Force switch to remove instances."
                    Write-Log -Message $ClassHasInstancesErr -Severity 2 -Source ${CmdletName} -DebugMessage
                    Write-Error -Message $ClassHasInstancesErr -Category 'InvalidOperation'
                }

                #  Create the ManagementClass Object
                [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:$ClassName")

                #  Remove the specified class properties
                $RemoveWmiProperty | ForEach-Object { $ClassObject.Properties.Remove($_) }

                #  Write the class and dispose of the object
                $null = $ClassObject.Put()
                $ClassObject.Dispose()
            }
            Else {
                $PropertyNotFoundErr = "No matching property [$Property] found for class [$Namespace`:$ClassName]."
                Write-Log -Message $PropertyNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $PropertyNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to remove property for class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {}
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Remove-WmiPropertyQualifier
Function Remove-WmiPropertyQualifier {
<#
.SYNOPSIS
    This function is used to remove WMI property qualifiers.
.DESCRIPTION
    This function is used to remove WMI class property qualifiers by name.
.PARAMETER Namespace
    Specifies the namespace. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name.
.PARAMETER PropertyName
    Specifies the property name for which to remove the qualifiers.
.PARAMETER QualifierName
    Specifies the property qualifier name or names.
.PARAMETER RemoveAll
    This switch is used to remove all qualifiers. Default is: $false. If this switch is specified the QualifierName parameter is ignored.
.PARAMETER Force
    This switch is used to remove all class instances. The class must be empty in order to be able to delete properties. Default is: $false.
.EXAMPLE
    Remove-WmiPropertyQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -PropertyName 'Source' -QualifierName 'Key','Description'
.EXAMPLE
    Remove-WmiPropertyQualifier -Namespace 'ROOT' -ClassName 'SCCMZone' -RemoveAll -Force
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyName,
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [string[]]$QualifierName,
        [Parameter(Mandatory=$false,Position=4)]
        [ValidateNotNullorEmpty()]
        [switch]$RemoveAll = $false,
        [Parameter(Mandatory=$false,Position=5)]
        [ValidateNotNullorEmpty()]
        [switch]$Force = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Get property qualifier names
            [string[]]$WmiPropertyQualifierNames = (Get-WmiPropertyQualifier -Namespace $Namespace -ClassName $ClassName -PropertyName $PropertyName -ErrorAction 'Stop').Name

            ## Get class instances
            $InstanceTest = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction 'SilentlyContinue'

            ## Add property qualifiers to deletion string array depending on selected options
            If ($RemoveAll) {
                $RemovePropertyQualifier = $ClassPropertyQualifierNames
            }
            ElseIf ($QualifierName) {
                $RemovePropertyQualifier = $WmiPropertyQualifierNames | Where-Object { $_ -in $QualifierName }
            }
            Else {
                $QualifierNameIsNullErr = "QualifierName cannot be `$null if -RemoveAll is not specified."
                Write-Log -Message $QualifierNameIsNullErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $QualifierNameIsNullErr -Category 'InvalidArgument'
            }

            ## Remove property qualifiers
            If ($RemovePropertyQualifier) {

                #  Remove all existing instances if the -Force switch was specified
                If ($Force -and $InstanceTest) {
                    Remove-WmiInstance -Namespace $Namespace -ClassName $ClassName -RemoveAll -ErrorAction 'Stop'
                }
                ElseIf ($InstanceTest) {
                    $ClassHasInstancesErr = "Instances [$($InstanceTest.Count)] detected in class [$Namespace`:$ClassName]. Use the -Force switch to remove instances."
                    Write-Log -Message $ClassHasInstancesErr -Severity 2 -Source ${CmdletName} -DebugMessage
                    Write-Error -Message $ClassHasInstancesErr -Category 'InvalidOperation'
                }

                #  Create the ManagementClass Object
                [wmiclass]$ClassObject = New-Object -TypeName 'System.Management.ManagementClass' -ArgumentList @("\\.\$Namespace`:$ClassName")

                #  Remove the specified property qualifiers
                $RemovePropertyQualifier | ForEach-Object { $ClassObject.Properties[$Property].Qualifiers.Remove($_) }

                #  Write the class and dispose of the object
                $null = $ClassObject.Put()
                $ClassObject.Dispose()
            }
            Else {
                $ProperyQualifierNotFoundErr = "No matching property qualifier [$Property`($QualifierName`)] found for class [$Namespace`:$ClassName]."
                Write-Log -Message $ProperyQualifierNotFoundErr -Severity 2 -Source ${CmdletName}
                Write-Error -Message $ProperyQualifierNotFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to remove property qualifier for class [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {}
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Remove-WmiInstance
Function Remove-WmiInstance {
<#
.SYNOPSIS
    This function is used to remove one ore more WMI instances.
.DESCRIPTION
    This function is used to remove one ore more WMI class instances with the specified values using CIM.
.PARAMETER Namespace
    Specifies the namespace where to search for the WMI namespace. Default is: 'ROOT\cimv2'.
.PARAMETER ClassName
    Specifies the class name from which to remove the instances.
.PARAMETER Property
    The class instance property to match. Can be piped. If there is more than one matching instance and the RemoveAll switch is not specified, an error will be thrown. 
.PARAMETER RemoveAll
    Removes all matching or existing instances.
.EXAMPLE
    [hashtable]$Property = @{
        'ServerPort' = '80'
        'ServerIP' = '10.10.10.11'
    }
    Remove-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Property $Property -RemoveAll
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,ValueFromPipeline,Position=2)]
        [ValidateNotNullorEmpty()]
        [hashtable]$Property,
        [Parameter(Mandatory=$false,Position=3)]
        [switch]$RemoveAll
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Get all class instances. If the class has no instances an error will be thrown
            $WmiInstances = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Stop'

            ## If Property was specified check for matching instances, otherwise if -RemoveAll switch is specified tag all instances for deletion
            If ($Property) {
                $RemoveInstances = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName -Property $Property -ErrorAction 'SilentlyContinue'
            }
            Else {
                $RemoveInstances = $WmiInstances
            }

            ## Remove according to specified options. If multiple instances are found check for the -RemoveAll switch
            If (($RemoveInstances.Count -eq 1) -or (($RemoveInstances.Count -gt 1) -and $RemoveAll)) {
                #  Remove instances one by one
                $RemoveInstances | ForEach-Object { Remove-CimInstance -InputObject $_ -ErrorAction 'Stop' }
            }
            
            ## Otherwise if more than one instance is detected, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            ElseIf ($RemoveInstances.Count -gt 1) {
                $MultipleInstancesFoundErr  = "Failed to remove instance. Multiple instances [$($RemoveInstances.Count)] found in class [$Namespace`:$ClassName]."
                Write-Log -Message $MultipleInstancesFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $MultipleInstancesFoundErr -Category 'InvalidOperation'
            }

            ## On instance removal failure, write debug message and optionally throw error if -ErrorAction 'Stop' is specified
            ElseIf (-not $RemoveInstances) {
                $InstanceNotFoundErr = "Failed to remove instances. No instances (or matching) found in class [$Namespace`:$ClassName]."
                Write-Log -Message $InstanceNotFoundErr -Severity 2 -Source ${CmdletName} -DebugMessage
                Write-Error -Message $MultipleInstancesFoundErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to remove instances in [$Namespace`:$ClassName]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $RemoveInstances
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Copy-WmiClass
Function Copy-WmiClass {
<#
.SYNOPSIS
    This function is used to copy a WMI class.
.DESCRIPTION
    This function is used to copy a WMI class to another namespace.
.PARAMETER ClassName
    Specifies the class name to be copied.
.PARAMETER NamespaceSource
    Specifies the source namespace where to search for the source WMI class. Default is: 'ROOT\cimv2'.
.PARAMETER NamespaceDestination
    Specifies the destination namespace.
.PARAMETER CreateDestination
    This switch is used to create the destination namespace if it does not exist. Default is: $false.
.EXAMPLE
    Copy-WmiClass -ClassName 'SCCMZone' -NamespaceSource 'ROOT' -NamespaceDestination 'ROOT\SCCMZone' -CreateDestination
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$NamespaceSource = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$NamespaceDestination,
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [switch]$CreateDestination = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the class exists in the source location
            $null = Get-WmiClass -Namespace $NamespaceSource -ClassName $ClassName -ErrorAction 'Stop'

            ## Get source class instances
            $InstanceTest = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction 'SilentlyContinue'

            ## Check if the destination namespace exists
            $DestinationTest = Get-WmiNameSpace -Namespace $NamespaceDestination -ErrorAction 'SilentlyContinue'

            ## Create destination namespace if specified
            If ($CreateDestination -and (-not $DestinationTest)) {

                #  Create destination namespace
                $null = New-WmiNameSpace -Namespace $NamespaceDestination -ErrorAction 'Continue'
            }
            ElseIf (-not $DestinationTest) {
                $DestinationNotFoundErr = "Destination namespace [$NamespaceDestination] not found. Use -CreateDestination switch to create the destination automatically."
                Write-Log -Message $DestinationNotFoundErr -Severity 2 -Source ${CmdletName}
                Write-Error -Message $DestinationNotFoundErr -Category 'ObjectNotFound'
            }
            
            ## Copy class to destination namespace
            $CopyClass = (Get-WmiObject -ClassName $ClassName -Namespace $NamespaceSource -list).CopyTo($NamespaceDestination)

            ## Copy source class instances if any are found
            If ($InstanceTest) {
                $null = Copy-WmiInstance #####
            }
               
        }
        Catch {
            Write-Log -Message "Failed to copy class. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $CopyClass
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Copy-WmiInstance
Function Copy-WmiInstance {
<#
.SYNOPSIS
    This function is used to copy the instances of a WMI class.
.DESCRIPTION
    This function is used to copy the instances of a WMI class to another class.
.PARAMETER ClassSourcePath
    Specifies the class to be copied from.
.PARAMETER ClassDestinationPath
    Specifies the class to be copied to.
.PARAMETER CreateDestination
    This switch is used to create the destination if it does not exist. Default is: $false.
.EXAMPLE
    Copy-WmiInstance -ClassSource 'ROOT\SCCM:SCCMZone' -ClassDestination 'ROOT\SCCM:SCCMZoneBlog' -CreateDestination
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$ClassSourcePath,
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassDestinationPath,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [switch]$CreateDestination = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {
            
            ## Set source and destination paths and name variables
            #  Set NamespaceSource
            $NamespaceSource = (Split-Path -Path $ClassSourcePath -Qualifier).TrimEnd(':')
            #  Set NamespaceDestination
            $NamespaceDestination =  (Split-Path -Path $ClassDestinationPath -Qualifier).TrimEnd(':')
            #  Set ClassNameSource
            $ClassNameSource = (Split-Path -Path $ClassSourcePath -NoQualifier)
            #  Set ClassNameDestination
            $ClassNameDestination = (Split-Path -Path $ClassDestinationPath -NoQualifier)

            ## Check if the source class exists. If source class does not exist throw an error
            $null = Get-WmiClass -Namespace $NamespaceSource -ClassName $ClassNameSource -ErrorAction 'Stop'

            ## Get the class source properties
            $ClassPropertiesSource = Get-WmiProperty -Namespace $NamespaceSource -ClassName $ClassNameSource -ErrorAction 'SilentlyContinue'

            ## Check if the destination class exists
            $ClassDestinationTest = Get-WmiClass -Namespace $NamespaceDestination -ClassName $ClassNameDestination -ErrorAction 'SilentlyContinue'

            ## Create destination class if specified
            If ((-not $ClassDestinationTest) -and $CreateDestination) {
                $null = Copy-WmiClassQualifier -ClassSourcePath $ClassSourcePath -ClassDestinationPath $ClassDestinationPath -CreateDestination -ErrorAction 'Stop'

                ## Get destination class properties
                $ClassPropertiesDestination = Get-WmiProperty -Namespace $NamespaceDestination -ClassName $ClassNameDestination -ErrorAction 'SilentlyContinue'

                ## Copy class properties from the source class if not present in the destination class
                $ClassPropertiesSource | ForEach-Object {
                    If ($_.Name -notin $ClassPropertiesDestination.Name) {
                        #  Create property
                        $null = New-WmiProperty -Namespace $NamespaceDestination -ClassName $ClassNameDestination -PropertyName $_.Name -PropertyType $_.CimType
                        #  Set qualifier if present
                        If ($_.Qualifiers.Name) {
                            $null = Set-WmiPropertyQualifier -Namespace $NamespaceDestination -ClassName $ClassNameDestination -PropertyName $_.Name -Qualifier @{ Name = $_.Qualifiers.Name; Value = $_.Qualifiers.Value }
                        }
                    }
                }

                ## Get source class instances
                $ClassInstancesSource =  Get-WmiInstance -Namespace $NamespaceSource -ClassName $ClassNameSource -ErrorAction 'SilentlyContinue' | Select-Object -Property $ClassPropertiesSource.Name

                ## Copy instances if thery are present in the source class ignoring any errors
                If ($ClassInstancesSource) {

                    #  Convert instance to hashtable
                    $ClassInstancesSource | ForEach-Object {

                        #  Initialize/Reset $InstanceProperty hashtable
                        $InstanceProperty = @{}

                        #  Assemble instance property hashtable
                        For ($i = 0; $i -le $($ClassPropertiesSource.Name.Length -1); $i++) {
                            $InstanceProperty += [ordered]@{
                                $($ClassPropertiesSource.Name[$i]) = $_.($ClassPropertiesSource.Name[$i])
                            }
                        }

                        #  Check if instance already in destination class
                        $ClassInstanceTest = [boolean]$(Get-WmiInstance -Namespace $NamespaceDestination -ClassName $ClassNameDestination -Property $InstanceProperty -ErrorAction 'SilentlyContinue')
                        
                        #  Create instance
                        If (-not $ClassInstanceTest) {
                            New-WmiInstance -Namespace $NamespaceDestination -ClassName $ClassNameDestination -Property $InstanceProperty -ErrorAction 'Stop'
                        }
                        Else {
                            Write-Log -Message "Instance already in destination class [$NamespaceDestination`:$ClassNameDestination]." -Severity 2 -Source ${CmdletName} -DebugMessage
                        }
                    }
                }
                Else {
                    Write-Log -Message  "No instances found in source class [$NamespaceSource`:$ClassNameSource]." -Severity 2 -Source ${CmdletName} -DebugMessage
                }
            }
            ElseIf (-not $ClassDestinationTest) {
                $DestinationClassErr = "Destination [$NamespaceDestination`:$ClassNameDestination] does not exist. Use the -CreateDestination switch to automatically create the destination class."
                Write-Log -Message $DestinationClassErr -Severity 2 -Source ${CmdletName}
                Write-Error -Message $DestinationClassErr -Category 'ObjectNotFound'
            }
        }
        Catch {
            Write-Log -Message "Failed to copy class instances. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
        }
        Finally {
            Write-Output -InputObject $CopyInstancesOutput
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Copy-WmiClassQualifier
Function Copy-WmiClassQualifier {
<#
.SYNOPSIS
    This function is used to copy the qualifiers of a WMI class.
.DESCRIPTION
    This function is used to copy the qualifiers of a WMI class to another class. Default qualifier flavors will be used.
.PARAMETER ClassSourcePath
    Specifies the class to be copied from.
.PARAMETER ClassDestinationPath
    Specifies the class to be copied to.
.PARAMETER QualifierName
    Specifies the class qualifier name or names to copy. Default is: 'All'.
.PARAMETER CreateDestination
    This switch is used to create the destination if it does not exist. Default is: $false.
.EXAMPLE
    Copy-WmiClassQualifier -ClassSourcePath 'ROOT\SCCM:SCCMZone' -ClassDestinationPath 'ROOT\SCCM:SCCMZoneBlog' -CreateDestination
.EXAMPLE
    Copy-WmiClassQualifier -ClassSourcePath 'ROOT\SCCM:SCCMZone' -ClassDestinationPath 'ROOT\SCCM:SCCMZoneBlog' -QualifierName 'Description' -CreateDestination
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$ClassSourcePath,
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassDestinationPath,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string[]]$QualifierName = 'All',
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [switch]$CreateDestination = $false
    )

    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Set source and destination paths and name variables
            #  Set NamespaceSource
            $NamespaceSource = (Split-Path -Path $ClassSourcePath -Qualifier).TrimEnd(':')
            #  Set NamespaceDestination
            $NamespaceDestination =  (Split-Path -Path $ClassDestinationPath -Qualifier).TrimEnd(':')
            #  Set ClassNameSource
            $ClassNameSource = (Split-Path -Path $ClassSourcePath -NoQualifier)
            #  Set ClassNameDestination
            $ClassNameDestination = (Split-Path -Path $ClassDestinationPath -NoQualifier)

            ## Check if source class exists
            $null = Get-WmiClass -Namespace $NamespaceSource -ClassName $ClassNameSource -ErrorAction 'Stop'

            ## Get source class qualifiers
            $ClassQualifiersSource = Get-WmiClassQualifier -Namespace $NamespaceSource -ClassName $ClassNameSource -ErrorAction 'SilentlyContinue'

            ## Check if the destination class exists
            $ClassDestinationTest = Get-WmiClass -Namespace $NamespaceDestination -ClassName $ClassNameDestination -ErrorAction 'SilentlyContinue'

            ## Create destination namespace and class if specified
            If ((-not $ClassDestinationTest) -and $CreateDestination) {
                $null = New-WmiClass -Namespace $NamespaceDestination -ClassName $ClassNameDestination -CreateDestination -ErrorAction 'Stop'
            }
            ElseIf (-not $ClassDestinationTest) {
                $DestinationClassErr = "Destination [$NamespaceSource`:$ClassName] does not exist. Use the -CreateDestination switch to automatically create the destination class."
                Write-Log -Message $DestinationClassErr -Severity 2 -Source ${CmdletName}
                Write-Error -Message $DestinationClassErr -Category 'ObjectNotFound'
            }

            ## Check if there are any qualifers in the source class 
            If ($ClassQualifiersSource) {
                
                ## Copy all qualifiers if not specified otherwise
                If ('All' -eq $QualifierName) {
                
                    #  Set destination class qualifiers 
                    $ClassQualifiersSource | ForEach-Object {
                        #  Set class qualifiers one by one
                        $CopyClassQualifier = Set-WmiClassQualifier -Namespace $NamespaceDestination -ClassName $ClassNameDestination -Qualifier @{ Name = $_.Name; Value = $_.Value } -ErrorAction 'Stop'
                    }
                }
                Else {

                    ## Copy class qualifier if it exists in source class, otherwise log the error and continue 
                    $ClassQualifiersSource | ForEach-Object {
                        If ($_.Name -in $QualifierName) {
                            $CopyClassQualifier = Set-WmiClassQualifier -Namespace $NamespaceDestination -ClassName $ClassNameDestination -Qualifier @{ Name = $_.Name; Value = $_.Value } -ErrorAction 'Stop'
                        }
                        Else {
                            $ClassQualifierNotFoundErr = "Failed to copy class qualifier [$($_.Name)]. Qualifier not found in source class [$NamespaceSource`:$ClassName]."
                            Write-Log -Message $ClassQualifierNotFoundErr -Severity 3 -Source ${CmdletName}
                        }
                    }
                }
            }
            Else {

                ## If no class qualifiers are found log error but continue execution regardless of the $ErrorActionPreference variable value
                Write-Log -Message "No qualifiers found in source class [$NamespaceSource`:$ClassName]." -Severity 2 -Source ${CmdletName}
            }
        }
        Catch {
            Write-Log -Message "Failed to copy class qualifier. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $CopyClassQualifier
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Copy-WmiProperty
Function Copy-WmiProperty {
<#
.SYNOPSIS
    This function is used to copy the properties of a WMI class.
.DESCRIPTION
    This function is used to copy the properties of a WMI class to another class. Default qualifier flavors will be used.
.PARAMETER ClassSourcePath
    Specifies the class to be copied from.
.PARAMETER ClassDestinationPath
    Specifies the class to be copied to.
.PARAMETER PropertyName
    Specifies the property name or names to copy. Default is: 'All'.
.PARAMETER CreateDestination
    This switch is used to create the destination if it does not exist. Default is: $false.
.EXAMPLE
    Copy-WmiProperty -ClassSourcePath 'ROOT\SCCM:SCCMZone' -ClassDestinationPath 'ROOT\SCCM:SCCMZoneBlog' -CreateDestination
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$ClassSourcePath,
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassDestinationPath,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$PropertyName = 'All',
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullorEmpty()]
        [switch]$CreateDestination = $false
    )
    
    Begin {
        ## Get the name of this function and write header
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Set source and destination paths and name variables
            #  Set NamespaceSource
            $NamespaceSource = (Split-Path -Path $ClassSourcePath -Qualifier).TrimEnd(':')
            #  Set NamespaceDestination
            $NamespaceDestination =  (Split-Path -Path $ClassDestinationPath -Qualifier).TrimEnd(':')
            #  Set ClassNameSource
            $ClassNameSource = (Split-Path -Path $ClassSourcePath -NoQualifier)
            #  Set ClassNameDestination
            $ClassNameDestination = (Split-Path -Path $ClassDestinationPath -NoQualifier)

            ## Check if source class exists
            $null = Get-WmiClass -Namespace $NamespaceSource -ClassName $ClassNameSource -ErrorAction 'Stop'

            ## Get source class properties
            $ClassPropertiesSource = Get-WmiProperty -Namespace $NamespaceSource -ClassName $ClassNameSource -ErrorAction 'SilentlyContinue'

            ## Check if the destination class exists
            $ClassDestinationTest = Get-WmiClass -Namespace $NamespaceDestination -ClassName $ClassNameDestination -ErrorAction 'SilentlyContinue'

            ## Create destination class if specified
            If ((-not $ClassDestinationTest) -and $CreateDestination) {
                $null = Copy-WmiClassQualifier -ClassSourcePath $ClassSourcePath -ClassDestinationPath $ClassDestinationPath -CreateDestination -ErrorAction 'Stop'
            }
            ElseIf (-not $ClassDestinationTest) {
                $DestinationClassErr = "Destination [$NamespaceSource`:$ClassName] does not exist. Use the -CreateDestination switch to automatically create the destination class."
                Write-Log -Message $DestinationClassErr -Severity 2 -Source ${CmdletName}
                Write-Error -Message $DestinationClassErr -Category 'ObjectNotFound'
            }

            ## Check if there are any properties in if not specified otherwiser
            If ($ClassPropertiesSource) {       

                ## Copy all class properties if not specified otherwise
                If ('All' -eq $PropertyName) {
                    
                    #  Create destination property and property qualifiers one by one
                    $ClassPropertiesSource | ForEach-Object {
                        #  Create property
                        $CopyClassProperty = New-WmiProperty -Namespace $NamespaceDestination -ClassName $ClassNameDestination -PropertyName $_.Name -PropertyType $_.CimType
                        #  Set qualifier if present
                        If ($_.Qualifiers.Name) {
                            $null = Set-WmiPropertyQualifier -Namespace $NamespaceDestination -ClassName $ClassNameDestination -PropertyName $_.Name -Qualifier @{ Name = $_.Qualifiers.Name; Value = $_.Qualifiers.Value }
                        }
                    }
                }
                Else {
                    
                    ## Copy specified property and property qualifier if it exists in source class, otherwise log the error and continue 
                    $ClassPropertiesSource | ForEach-Object {
                        If ($_.Name -in $PropertyName) {
                            #  Create property
                            $CopyClassProperty =  New-WmiProperty -Namespace $NamespaceDestination -ClassName $ClassNameDestination -PropertyName $_.Name -PropertyType $_.CimType
                            #  Set qualifier if present
                            If ($_.Qualifiers.Name) {
                                $null = Set-WmiPropertyQualifier -Namespace $NamespaceDestination -ClassName $ClassNameDestination -PropertyName $_.Name -Qualifier @{ Name = $_.Qualifiers.Name; Value = $_.Qualifiers.Value }
                            }
                        }
                        Else {
                            $ClassPropertyNotFoundErr = "Failed to copy class property [$($_.Name)]. Property not found in source class [$NamespaceSource`:$ClassName]."
                            Write-Log -Message $ClassPropertyNotFoundErr -Severity 3 -Source ${CmdletName}
                        }
                    }
                }
            }
            Else {

                ## If no class properties are found log error but continue execution regardless of the $ErrorActionPreference variable value
                Write-Log -Message "No properties found in source class [$NamespaceSource`:$ClassName]." -Severity 2 -Source ${CmdletName}
            }
        }
        Catch {
            Write-Log -Message "Failed to copy class property. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
            Break
        }
        Finally {
            Write-Output -InputObject $CopyClassProperty
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Copy-WmiNamespace
Function Copy-WmiNamespace {
<#
.SYNOPSIS
    This function is used to copy a WMI namespace.
.DESCRIPTION
    This function is used to copy a WMI namespace to another namespace.
.PARAMETER NamespaceSource
    Specifies the source namespace where to search for the source WMI namespace.
.PARAMETER NamespaceDestination
    Specifies the destination namespace. Default is: ROOT\cimv2.
.PARAMETER CreateDestination
    This switch is used to create the destination namespace if it does not exist. Default is: $false.
.EXAMPLE
    Copy-WmiNamespace -NamespaceSource 'ROOT\SCCMZone' -NamespaceDestination 'ROOT\cimv2'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$NamespaceSource,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$NamespaceDestination = 'ROOT\cimv2',
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [switch]$CreateDestination = $false
    )

    Begin {
    ## Get the name of this function and write header
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
    Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the source namespace exists
            $SourceTest = Get-WmiNameSpace -Namespace $NamespaceSource

            ## Check if the destination namespace exists
            $DestinationTest = Get-WmiNameSpace -Namespace $NamespaceDestination

            ## If the source exists
            If ($SourceTest) {

                ## Create CopyNamespaceScriptBlock
                [scriptblock]$CopyNamespaceScriptBlock = {

                    #  Set Destination Namespace Name
                    [string]$NamespaceName = Split-Path $NamespaceSource -Leaf
                    #  Set Destination Namespace
                    [string]$Namespace = Join-Path -Path $NamespaceDestination -ChildPath $NamespaceName

                    #  Create Destination Namespace
                    New-WmiNameSpace -Namespace $NamespaceDestination -NamespaceName $NamespaceName

                    #  Copy classes from Source to Destination
                    Copy-WmiClass -NamespaceSource $RootNamespaceSource -NamespaceDestination $RootNamespace
                }

                ## If the Destination exists execute $CopyNamespaceScriptBlock
                If ($DestinationTest) { $CopyNamespace = & $CopyNamespaceScriptBlock }

                ## If the -CreateDestination switch was specified, create the Destination Root Namespace and execute $CopyClassScriptBlock
                ElseIf ($CreateDestination) {

                    #  Set Root Destination Namespace
                    [string]$RootNamespaceName = Split-Path -Path $NamespaceDestination -leaf
                    #  Set Root Destination Namespace Name
                    [string]$RootNamespaceName = Split-Path -Path $NamespaceDestination -leaf

                    #  Create Root Destination Namespace
                    New-WmiNameSpace -Namespace $Namespace -NamespaceName $NamespaceName

                    #  Copy Namespace
                    $CopyNamespace = & $CopyClassScriptBlock
                }
                Else {
                    Write-Log -Message "Failed to copy namespace. $NamespaceDestination (destination) does not exist." -Severity 3 -Source ${CmdletName}
                }
            }
            Else {
                Write-Log -Message "Failed to copy namespace. $NamespaceSource (source) does not exist." -Severity 3 -Source ${CmdletName}
            }
        }
        Catch {
            Write-Log -Message "Failed to copy namespace. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
        }
        Finally {
            Write-Output -InputObject $CopyNamespace
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Rename-WmiNamespace
Function Rename-WmiNamespace {
<#
.SYNOPSIS
    This function is used to rename a WMI namespace.
.DESCRIPTION
    This function is used to rename a WMI namespace by creating a new namespace, copying all existing classes to it and removing the old one.
.PARAMETER Namespace
    Specifies the root namespace where to search for the namespace name. Default is: ROOT\cimv2.
.PARAMETER NamespaceName
    Specifies the namespace name to be renamed.
.PARAMETER NamespaceNewName
    Specifies the new namespace name.
.EXAMPLE
    Rename-WmiNamespace -Namespace 'ROOT\cimv2' -NamespaceName 'OldName' -NamespaceNewName 'NewName'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$NamespaceName,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$NamespaceNewName
    )

    Begin {
    ## Get the name of this function and write header
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
    Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
    }
    Process {
        Try {

            ## Check if the namespace to be renamed exists
            $NamespaceTest = Get-WmiNameSpace -Namespace $Namespace -NamespaceName $NamespaceName

            ## Check if the new namespace exists
            $NewNamespaceNameTest = Get-WmiNameSpace -Namespace $Namespace -NamespaceName $NamespaceNewName

            ## If the namespace exists and another namespace with the new name does not exist, rename namespace
            If ($NamespaceTest) {
                If (-not $NewNamespaceNameTest) {

                    #  Create New Namespace
                    New-WmiNameSpace -Namespace $Namespace -NamespaceName $NamespaceName

                    #  Set Source Namespace
                    $NamespaceSource = Join-Path -Path $Namespace -ChildPath $NamespaceName

                    #  Set Destination Namespace
                    $NamespaceDestination = Join-Path -Path $Namespace -ChildPath $NamespaceNewName

                    #  Copy clases to the new Namespace
                    Copy-WmiClass -NamespaceSource $NamespaceSource -NamespaceDestination $NamespaceDestination

                    #  Remove old Namespace
                    Remove-WmiNameSpace -Namespace $Namespace -NamespaceName $NamespaceName

                    #  Write success message to console
                    Write-Log -Message "Succesfully renamed $Namespace`\$NamespaceName to $Namespace`\$NamespaceNewName." -Source ${CmdletName}
                }
                Else {
                    Write-Log -Message "Failed to rename namespace. $Namespace`\$NamespaceNewName already exists." -Severity 3 -Source ${CmdletName}
                }
            }
            Else {
                Write-Log -Message "Failed to rename namespace. $Namespace`\$NamespaceName does not exist." -Severity 3 -Source ${CmdletName}
            }
        }
        Catch {
            Write-Log -Message "Failed to rename namespace. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
        }
        Finally {}
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Rename-WmiClass
Function Rename-WmiClass {
<#
.SYNOPSIS
    This function is used to rename a WMI class.
.DESCRIPTION
    This function is used to rename a WMI class by creating a new class, copying all existing properties and instances to the new class, and removing the old one.
.PARAMETER Namespace
    Specifies the namespace where the class is located. Default is: ROOT\cimv2.
.PARAMETER ClassName
    Specifies the name of the class to be renamed.
.PARAMETER ClassNewName
    Specifies the new class name.
.EXAMPLE
    Rename-WmiClass -Namespace 'Root\cimv2' -ClassName 'OldName' -ClassNewName 'NewName'
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$ClassName,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$ClassNewName
    )

    Begin {
    ## Get the name of this function and write header
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header

        ## Set Connection Props
        #  Set class Connection Props
        [hashtable]$ClassConnectionProps = @{ NameSpace = $Namespace; ClassName = $ClassName }
        #  Set new class Connection Props
        [hashtable]$NewClassConnectionProps = @{ NameSpace = $Namespace; ClassName = $ClassNewName }
    }
    Process {
        Try {

            [hashtable]$ClassConnectionProps = @{NameSpace = 'ROOT\TEST'; ClassName = 'Tivo'}

            ## Check if the class to be renamed exists
            $ClassNameTest = Get-WmiClass @ClassConnectionProps  -ErrorAction 'SilentlyContinue'

            ## Check if the new class exists
            $ClassNameTest = Get-WmiClass @NewClassConnectionProps  -ErrorAction 'SilentlyContinue'

            ## If the class exists and another class with the new name does not exist, rename class
            If ($ClassNameTest) {

            }
                If (-not $ClassNewNameTest) {




                #  Get Old Class instances
                $OldInstances = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName

                $OldInstances.CopyTo()
                #  Remove old Class
                Remove-WmiClass -Namespace $Namespace -ClassName $ClassName

                #  Return Result
                $Result = "Rename Class $Namespace`:$ClassName to $Namespace`:$ClassNewName - Success!"
                }
                Else {
                    $Result = "Rename Class - Failed! $Namespace`:$ClassNewName already exists!"
                }

        }
        Catch {
            $Result = "Rename Class - Failed! `n $_"
        }
        Finally {
            Write-Output -InputObject $Result
        }
    }
    End {
        Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
    }
}
#endregion


#region Function Set-WmiInstance
Function Set-WmiInstance {
<#
.SYNOPSIS
    This function is used to modify a WMI Instance.
.DESCRIPTION
    This function is used to modify or optionaly creating a WMI Instance if it does not exist using CIM.
.PARAMETER Namespace
    The Class Namespace.
    The Default is ROOT\cimv2.
.PARAMETER ClassName
    The Class Name.
.PARAMETER Key
    The Properties that are used as keys (Optional).
.PARAMETER PropertySearch
    The Class Instance Properties and Values to find.
.PARAMETER Property
    The Class Instance Properties and Values to set.
.PARAMETER CreateInstance
    Switch for creating the instance if it does not exist.
    Default is $false
.EXAMPLE
    [hashtable]$PropertySearch = @{
        'ServerPort' = '99'
        'ServerIP' = '10.10.10.10'
    }
    [hashtable]$Property = @{
        'ServerPort' = '88'
        'ServerIP' = '11.11.11.11'
        'Source' = 'File1'
        'Date' = $(Get-Date)
    }
    Set-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Key 'File1' -PropertySearch $PropertySearch -Property $Property
    Set-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Key 'File1' -Property $Property
    Set-WmiInstance -Namespace 'ROOT' -ClassName 'SCCMZone' -Property $Property -CreateInstance
.NOTES
    This is a module function and can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [string]$Namespace = 'ROOT\cimv2',
        [Parameter(Mandatory=$false,Position=1)]
        [string]$ClassName = 'SCCMZone',
        [Parameter(Mandatory=$false,Position=2)]
        [string[]]$Key = '',
        [Parameter(Mandatory=$false,Position=3)]
        [hashtable]$PropertySearch = '',
        [Parameter(Mandatory=$true,Position=4)]
        [hashtable]$Property,
        [Parameter(Mandatory=$false,Position=5)]
        [switch]$CreateInstance = $false,
        [PSCustomObject]$Result = @()
    )

    Try {

                #move to get instance?

                #  Get Property Names from function input to be used for filtering
                [string[]]$ClassPropertyNames =  $($Property.GetEnumerator().Name)

                #  Get all Instances for the specified Wmi Class, selecting only specified Property Names
                [PSCustomObject]$ClassInstances = Get-CimInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction 'Continue' | Select-Object $ClassPropertyNames

                #  Convert Property hashtable to PSCustomObject for comparison
                [PSCustomObject]$InputProperty = [PSCustomObject]$Property

                #  -ErrorAction 'SilentlyContinue' does not seem to work correctly with the Compare-Object commandlet so it needs to be set globaly
                $ErrorActionPreferenceOriginal = $ErrorActionPreference
                $ErrorActionPreference = 'SilentlyContinue'

                #  Check if and instance with the same values exists. Since $InputProperty is a dinamically generated object Compare-Object has no hope of working correctly.
                #  Luckily Compare-Object as a -Property parameter which allows us to look at specific parameters.
                $InstanceSearch = Compare-Object -ReferenceObject $InputProperty -DifferenceObject $ClassInstances -Property $ClassPropertyNames -IncludeEqual -ExcludeDifferent

                #  Setting the ErrorActionPreference back to the previous value
                $ErrorActionPreference = $ErrorActionPreferenceOriginal

                #  If no matching instance is found, create a new instance, else write error
                If (-not $InstanceSearch) {
                    $NewInstance = & $NewInstanceScriptBlock
                }




        ## Set Connection Props
        [hashtable]$ConnectionProps = @{ NameSpace = $Namespace; ClassName = $ClassName }

        ## Test if the Class exists
        [bool]$ClassTest = Get-WmiClass -Namespace $Namespace -ClassName $ClassName
        If ($ClassTest) {

            ## If -PropertySearch parameter was specified use it to get the instances
            If ($PropertySearch) {
                $InstanceTest = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName -Property $InputProperty -ErrorAction 'SilentlyContinue'
            }

            ## If the -PropertySearch parameter was not specified, use the -Property do get the instances
            Else {
                $InstanceTest = Get-WmiInstance -Namespace $Namespace -ClassName $ClassName -Property $Property -ErrorAction 'SilentlyContinue'
            }

            ## Count Instances
            [int16]$InstanceCount = ($InstanceTest | Measure-Object).Count

            ## Perform actions depending on the $InstanceTest result
            Switch ($InstanceCount) {

                #  If $InstanceTest is not $null and contains just one instance, Set the new values
                '1' { $Result = $InstanceTest | Set-CimInstance -Property $Property -ErrorAction 'Stop' }

                #  If $InstanceTest is not $null and contains more than one instance, abort and return error message
                { $_ -gt '1' } { $Result = 'Set Instance - Failed! More than one instance with the specified values found!' }

                #  If $InstanceTest is $null, the -CreateInstance switch was specified and not matching instance exists, create a new instance with the specified values
                { $_ -eq '0' -and (-not $InstanceTest) -and $CreateInstance } {

                    #  Create a new instance with or without the key parameter
                    If ($Key) {
                        $Result = New-CimInstance -Namespace $Namespace -ClassName $ClassName -Key $Key -Property $Property -ErrorAction 'Stop'
                    }
                    Else {
                        $Result = New-CimInstance -Namespace $Namespace -ClassName $ClassName -Property $Property -ErrorAction 'Stop'
                    }
                }
                Default { $Result = 'Unhandled Exception!' }
            }
        }
        Else {
            $Result = "Set Instance - Failed! $Namespace`:$ClassName does not exist!"
        }
    }
    Catch {
        $Result = "Set Instance - Failed! `n $_"
    }
    Finally {
        Write-Output -InputObject $Result
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


Export-ModuleMember -Function Get-WmiNameSpace
Export-ModuleMember -Function Get-WmiClass
Export-ModuleMember -Function Get-WmiProperty
Export-ModuleMember -Function Get-WmiInstance

Export-ModuleMember -Function Get-WmiClassQualifier
Export-ModuleMember -Function Get-WmiPropertyQualifier

Export-ModuleMember -Function New-WmiNameSpace
Export-ModuleMember -Function New-WmiClass
Export-ModuleMember -Function New-WmiProperty
Export-ModuleMember -Function New-WmiInstance

Export-ModuleMember -Function Remove-WmiNameSpace
Export-ModuleMember -Function Remove-WmiClass
Export-ModuleMember -Function Remove-WmiProperty
Export-ModuleMember -Function Remove-WmiInstance

Export-ModuleMember -Function Remove-WmiClassQualifier
Export-ModuleMember -Function Remove-WmiPropertyQualifier

Export-ModuleMember -Function Copy-WmiNameSpace
Export-ModuleMember -Function Copy-WmiClass
Export-ModuleMember -Function Copy-WmiProperty
Export-ModuleMember -Function Copy-WmiInstance


Export-ModulMember -Function Rename-WmiNamespace
Export-ModulMember -Function Rename-WmiClass

Export-ModuleMember -Function Set-WmiClassQualifier
Export-ModuleMember -Function Set-WmiPropertyQualifier
Export-ModuleMember -Function Set-WmiInstance

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
