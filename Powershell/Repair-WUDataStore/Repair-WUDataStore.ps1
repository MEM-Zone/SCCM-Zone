<#
*********************************************************************************************************
* Requires          | Requires PowerShell 2.0                                                           *
* ===================================================================================================== *
* Modified by       |    Date    | Revision | Comments                                                  *
* _____________________________________________________________________________________________________ *
* Ioan Popovici     | 2018-03-28 | v1.0     | First version                                             *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    Repairs a corrupted WU DataStore.
.DESCRIPTION
    Detects and repairs a corrupted WU DataStore.
    Detection is done by counting ESENT errors from the last 3 days in the Application eventlog with the EventID equal to '623'.
    Repairs are performed by removing and reinitializing the corrupted DataStore.
    After the repair step completes the Application eventlog is backed up and cleared in order not to triger the detection again.
    The backup of the Application log is stored in 'SystemRoot\Temp' folder.
.PARAMETER Action
    Specifies the action to be performed. Available actions are: ('DetectAndRepair', 'Detect', 'Repair').
    Default is: 'DetectAndRepair'
.EXAMPLE
    Repair-WUDataStore -Action 'Detect'
.INPUTS
    System.String.
.OUTPUTS
    None. This function has no outputs.
.NOTES
    This function can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
.COMPONENT
    WindowsUpdate
.FUNCTIONALITY
    Repair
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

## Get script parameters
Param (
    [Parameter(Mandatory=$false,Position=0)]
    [ValidateNotNullorEmpty()]
    [ValidateSet('DetectAndRepair','Detect','Repair')]
    [string]$Action = 'DetectAndRepair'
)

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Backup-EventLog
Function Backup-EventLog {
<#
.SYNOPSIS
    Backs-up an Event Log.
.DESCRIPTION
    Backs-up an Event Log using the BackUpEventLog Cim method.
.PARAMETER LogName
    Specifies the event log to backup.
.PARAMETER BackupPath
    Specifies the Backup Path. Default is: '$Env:SystemRoot\Temp'.
.PARAMETER BackupName
    Specifies the Backup name. Default is: 'yyyy-MM-dd_HH-mm-ss_$Env:ComputerName_$LogName'.
.EXAMPLE
    Backup-EventLog -LogName 'Application' -BackupPath 'C:\SCCMZone' -BackupName '1980-09-09_10-10-00_SCCMZoneBlog_Application'
.INPUTS
    System.String.
.OUTPUTS
    None. This function has no outputs.
.NOTES
    This function can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
.COMPONENT
    EventLog
.FUNCTIONALITY
    Backup
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [string]$LogName,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [string]$BackupPath,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateNotNullorEmpty()]
        [string]$BackupName
    )

    Begin {

        ## Setting variables
        $PowerShellVersion = $PSVersionTable.PSVersion.Major
        $Date = $(Get-Date -f 'yyyy-MM-dd_HH-mm-ss')
        #  Setting optional parameters
        If (-not $BackupPath) {
            $BackupPath = $(Join-Path -Path $Env:SystemRoot -ChildPath '\Temp')
        }
        If (-not $BackupFileName) {
            $BackUpFileName = "{0}_{1}_{2}.evtx" -f $Date, $Env:COMPUTERNAME, $LogName
        }
        #  Setting backup arguments
        $BackupArguments = @{ ArchiveFileName = (Join-Path -Path $BackupPath -ChildPath $BackUpFileName) }
    }
    Process {
        Try {

            If ($PowerShellVersion -eq 2) {
                ## Get event log
                $EventLog = Get-WmiObject -Class 'Win32_NtEventLogFile' -Filter "LogFileName = '$LogName'"

                ## Backup event log
                $BackUp = $EventLog | Invoke-WmiMethod -Name 'BackupEventLog' -ArgumentList $BackupArguments -ErrorAction 'SilentlyContinue'

                # $BackUp = $EventLog | Invoke-CimMethod -Name 'BackupEventLog' -Arguments $BackupArguments -ErrorAction 'SilentlyContinue'
                If ($BackUp.ReturnValue -ne 0) {
                    Throw "Backup failed with $($BackUp.ReturnValue)"
                }
            }
            ElseIf ($PowerShellVersion -ge 3) {
                ## Get event log
                $EventLog = Get-CimInstance -ClassName 'Win32_NtEventLogFile' -Filter "LogFileName = '$LogName'"

                ## Backup event log
                $BackUp = $EventLog | Invoke-CimMethod -Name 'BackupEventLog' -Arguments $BackupArguments -ErrorAction 'SilentlyContinue'

                # $BackUp = $EventLog | Invoke-CimMethod -Name 'BackupEventLog' -Arguments $BackupArguments -ErrorAction 'SilentlyContinue'
                If ($BackUp.ReturnValue -ne 0) {
                    Throw "Backup failed with $($BackUp.ReturnValue)"
                }
            }
            Else {
                Throw "PowerShell version [$PowerShellVersion] not supported."
            }
        }
        Catch {
            Write-Output -InputObject "Failed to query EventLog [$LogName]. `n$_"
            Break
        }
    }
    End {
    }
}
#endregion

#region Function Test-EventLogCompliance
Function Test-EventLogCompliance {
<#
.SYNOPSIS
    Tests the EventLog compliance for specific events.
.DESCRIPTION
    Tests the EventLog compliance by getting events and returing a Non-Compliant statement after a specified treshold is reached.
.PARAMETER LogName
    Specifies the LogName to search.
.PARAMETER Source
    Specifies the Source to search.
.PARAMETER EventID
    Specifies the EventID to search.
.PARAMETER EntryType
    Specifies the Entry Type to search. Available options are: ('Information','Warning','Error'). Default is: 'Error'.
.PARAMETER After
    Specifies the data and time that this function get events that occur after. Enter a DateTime object, such as the one returned by the Get-Date cmdlet.
.PARAMETER Threshold
    Specifed the treshold after which this functions returns $true.
.EXAMPLE
    Test-EventLogCompliance -LogName 'Application' -Source 'ESENT' -EventID '623' -EntryType 'Error' -After $((Get-Date).AddDays(-1)) -Threshold 3
.INPUTS
    None.
.OUTPUTS
    System.Boolean.
.NOTES
    This function can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
.COMPONENT
    WindowsUpdate
.FUNCTIONALITY
    Test
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullorEmpty()]
        [String]$LogName,
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullorEmpty()]
        [String]$Source,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateNotNullorEmpty()]
        [String]$EventID,
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateSet('Information','Warning','Error')]
        [String]$EntryType = 'Error',
        [Parameter(Mandatory=$true,Position=4)]
        [ValidateNotNullorEmpty()]
        [DateTime]$After,
        [Parameter(Mandatory=$true,Position=5)]
        [ValidateNotNullorEmpty()]
        [Int16]$Threshold
    )

    Begin {

        ## Setting Variables
        $ErrorMessage = $null
    }
    Process {
        Try {
            $Events = Get-EventLog -LogName $LogName -Source $Source -EntryType $EntryType -After $After -ErrorAction 'Stop' | Where-Object { $_.EventID -eq $EventID }
        }
        Catch [System.ArgumentException] {

            $ErrorMessage = $_.Exception

            ## Continue if no matches are found
            If ($ErrorMessage -match 'No matches found') {
                Write-Output -InputObject 'Compliant'
                Continue
            }
            Else {
                Throw $ErrorMessage
            }
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            Throw $ErrorMessage
        }
        Finally {
            If ($Events.Count -ge $Threshold) {
                Write-Output -InputObject 'Non-Compliant'
            }
            ElseIf (-not $ErrorMessage) {
                Write-Output -InputObject 'Compliant'
            }
        }
    }
    End {
    }
}
#endregion

#region Function Repair-WUDataStore
Function Repair-WUDataStore {
<#
.SYNOPSIS
    Repairs a corrupted WU DataStore.
.DESCRIPTION
    Repairs a corrupted WU DataStore by removing and reinitializing the corrupted DataStore.
.EXAMPLE
    Repair-WUDataStore
.INPUTS
    None. This function has no inputs.
.OUTPUTS
    None. This function has no outputs.
.NOTES
    This function can typically be called directly.
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
.COMPONENT
    WindowsUpdate
.FUNCTIONALITY
    Repair
#>

    Begin {

        ## Setting Variables
        $ErrorMessage = $null
        #  Setting Paths
        $PathRegsvr = (Join-Path -Path $Env:SystemRoot -ChildPath '\System32\Regsvr32.exe')
        $PathDataStore = (Join-Path -Path $Env:SystemRoot -ChildPath '\SoftwareDistribution\DataStore')
    }
    Process {
        Try {

            ## Re-register wuauend.dll
            Start-Process -FilePath $PathRegsvr -ArgumentList '/s Wuaueng.dll' -Wait

            ## Stop the windows update service
            Stop-Service -Name 'wuauserv' -Force -ErrorAction 'SilentlyContinue'

            ## Wait for the windows update service to stop
            #  Setting Loop index to 12 (one minute)
            $Loop = 1
            While ($StatusWuaService -ne 'Stopped') {

                #  Waiting 5 seconds
                Start-Sleep -Seconds 5
                $StatusWuaService =  (Get-Service -Name 'wuauserv').Status

                #  Incrementing loop index
                $Loop++

                #  Throw error if service has not stopped within 1 minute
                If ($Loop -gt 7) {
                    Throw 'Timeout occured while waiting for Windows Update Service to stop'
                }
            }

            ## Remove the Windows update DataStore
            Remove-Item -Path $PathDataStore -Recurse -Force | Out-Null
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            Throw $ErrorMessage
        }
        Finally {
            If (-not $ErrorMessage) {
                Write-Output -InputObject 'Remediated'

            }
        }
    }
    End {
        Try {

            ## Start the windows update service
            Start-Service -Name 'wuauserv' -ErrorAction 'Stop'
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            Throw $ErrorMessage
        }
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

Switch ($Action) {
    'DetectAndRepair' {
        $ESENTError623 = Test-EventLogCompliance -LogName 'Application' -Source 'ESENT' -EventID '623' -EntryType 'Error' -After $((Get-Date).AddDays(-3)) -Threshold 3
        Write-Output -InputObject $ESENTError623

        If ($ESENTError623 -eq 'Non-Compliant') {
            Repair-WUDataStore
            Backup-EventLog -LogName 'Application'
            Clear-EventLog -LogName 'Application'
        }
    }
    'Detect' {
        $ESENTError623 = Test-EventLogCompliance -LogName 'Application' -Source 'ESENT' -EventID '623' -EntryType 'Error' -After $((Get-Date).AddDays(-3)) -Threshold 3
        Write-Output -InputObject $ESENTError623
    }
    'Repair' {
        Write-Output -InputObject 'Reparing without checking compliance'
        Repair-WUDataStore
        Backup-EventLog -LogName 'Application'
        Clear-EventLog -LogName 'Application'
    }
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================