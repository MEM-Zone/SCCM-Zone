<#
*********************************************************************************************************
*                                                                                                       *
*** This PowerShell script is Get the SCEP Logs from SCCM DB and export them to a CSV file            ***
*                                                                                                       *
*********************************************************************************************************
* Created by Octavian Cordos, 2017-03-28   | Requirements: PowerShell 3.0, SQL PS commandlets           *
* ======================================================================================================*
* Modified by                   | Date       | Version  | Comments                                      *
*_______________________________________________________________________________________________________*
* Octavian Cordos               | 2017-03-28 | v1.0     | First version                                 *
* Ioan Popovici                 | 2017-03-28 | v1.1     | Minor modifications, cleanup                  *
*-------------------------------------------------------------------------------------------------------*
* Execute with: C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe -NoProfile -File              *
* Get-SCEPDetectionLogs.ps1 -CMSQLServer "SCCM_SQL_Server_Name" -CMDatabase "CM_Database_Name"          *
*********************************************************************************************************

    .SYNOPSIS
        This PowerShell script is Get the SCEP Logs from SCCM DB and export them to a CSV file
    .DESCRIPTION
        This PowerShell script is Get the SCEP Logs from SCCM Database and export them to a CSV file.
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

    ## External Script Variables
    Param (
        [Parameter(Mandatory=$True,Position=0)]
        [Alias('CMSQL')]
        [string]$CMSQLServer,
        [Parameter(Mandatory=$True,Position=1)]
        [Alias('CMDB')]
        [string]$CMDatabase
    )

    ## Get script path and name
    [String]$ScriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
    [String]$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)

    ## CSV file initialization
    #  Set the CSV Data file name
    [String]$csvDataFileName = $ScriptName
    [String]$csvDataFileNameWithExtension = $csvDataFileName+'.csv'

    #  Assemble CSV Data file path
    [String]$csvDataFilePath = (Join-Path -Path $ScriptPath -ChildPath $csvDataFileName)+'.csv'

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody

    ## Check if Get-SCEPDetectionLogs.cvs exists
    If (-not (Test-Path $csvDataFilePath)) {
        New-Item -Path $ScriptPath -Name $csvDataFileNameWithExtension -Type 'File'
    }

    ## Get Last Write Time Stamp
    $LastWriteTime = Import-CSV $csvDataFilePath | Select-Object -ExpandProperty Timestamp -Last 1

    ## Get Logs from Database using the CSV Last Write Time Stamp in order to pick up where the last Get-SCCMDetectionLog cycle ended
    $SCEPLogData = Invoke-SQLcmd -Query "SELECT * FROM $($CMDatabase).dbo.v_AM_NormalizedDetectionHistory SCEP WHERE SCEP.TimeStamp > `'$LastWriteTime`' ORDER BY SCEP.TimeStamp ASC" -ServerInstance $CMSQLServer

    ## Check if we have something to write to the CSV file
    If (-not ($SCEPLogData)) {

        # Overwrite the CSV file with gathered SCEP Log Data if the size gets over 2500 KB
        If ((Get-Item $csvDataFilePath).Length -gt 2500KB) {
            $SCEPLogData | Export-CSV $csvDataFilePath -NoTypeInformation -Force -Encoding 'UTF8' -ErrorAction 'Continue'
        }

        # Append gathered SCEP Log Data to the CSV file
        Else {
            $SCEPLogData | Export-CSV $csvDataFilePath -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Continue'
        }
    }

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
