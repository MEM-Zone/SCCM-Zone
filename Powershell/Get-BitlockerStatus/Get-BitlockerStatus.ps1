<#
*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 3.0                                                  *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2015-11-15 | v1.0     | First version                                                 *
* Ioan Popovici | 2017-09-22 | v1.1     | Modified for all drives, improvements and code cleanup        *
* Ioan Popovici | 2018-09-03 | v1.2     | Added encryption status                                       *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This PowerShell Script is used to get the BitLocker protection status.
.DESCRIPTION
    This PowerShell Script is used to get the BitLocker protection status for a specific drive, or all drives.
.EXAMPLE
    C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe -NoExit -NoProfile -File Get-BitLockerStatus.ps1
.LINK
    https://SCCM-Zone.com
    https://github.com/JhonnyTerminus/SCCMZone
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

## Initializing Result Object
[psCustomObject]$Result =@()
[array]$ResultProps =@()

## Initializing variables
[string]$LocalDrives

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Get-BitLockerStatus
Function Get-BitLockerStatus {
<#
.SYNOPSIS
    This Function is used the get the BitLocker Protection Status.
.DESCRIPTION
    This Function is used the get the BitLocker Protection Status.
.PARAMETER DriveLetter
    Drive Letter to check for BitLocker Status. Optional Parameter.
.EXAMPLE
    Get-BitLockerStatus -DriveLetter 'D'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://SCCM-Zone.com
    https://github.com/JhonnyTerminus/SCCMZone
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [Alias('Drive')]
        [string]$DriveLetter
    )

    Try {

        ##  Get the local drives from WMI
        $LocalDrives = Get-CimInstance -Namespace 'root\CIMV2' -ClassName 'CIM_LogicalDisk' | Where-Object { $_.DriveType -eq '3' }

        ## Get the BitLocker Status for all drives from WMI
        Get-CimInstance  -Namespace 'root\CIMV2\Security\MicrosoftVolumeEncryption'  -ClassName 'Win32_EncryptableVolume' -ErrorAction 'Stop' | ForEach-Object {

            #  Create the Result Props and make the ProtectionStatus more report friendly
            $ResultProps = [ordered]@{
                'Drive' = $_.DriveLetter
                'ProtectionStatus' = $(
                    Switch ($_.ProtectionStatus) {
                        0 { 'PROTECTION OFF' }
                        1 { 'PROTECTION ON' }
                        2 { 'PROTECTION UNKNOWN' }
                    }
                )
                'EncryptionStatus' = $(
                    Switch ($_.ConversionStatus) {
                        0 { 'FullyDecrypted' }
                        1 { 'FullyEncrypted' }
                        2 { 'EncryptionInProgress' }
                        3 { 'DecryptionInProgress' }
                        4 { 'EncryptionPaused' }
                        5 { 'DecryptionPaused' }
                    }
                )
            }

            #  Adding ResultProps hash table to result object
            $Result += New-Object PSObject -Property $ResultProps
        }

        #  Workaround for some Windows 7 computers not reporting BitLocker protection status for all drives
        #  Create the ResultProps array
        $LocalDrives | ForEach-Object {
            If ($_.DeviceID -notin $Result.Drive) {
                $ResultProps = [ordered]@{
                    'Drive' = $_.DeviceID
                    'ProtectionStatus' = 'PROTECTION OFF'
                }

                #  Adding ResultProps hash table to result object
                $Result += New-Object PSObject -Property $ResultProps
            }
        }
    }

    ## Catch any script errors
    Catch {
        Write-Host "Script Execution - Error!`n $_`n" -ForegroundColor 'Red' -BackgroundColor 'Black'
    }
    Finally {

        ## Return different Results depending on wether the DriveLetter parameter was specified or not
        If ($DriveLetter) {
            Write-Output -InputObject $($Result | Where-Object { $_.Drive -match $DriveLetter })
        }
        Else {
            Write-Output -InputObject $Result
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

## Write BitLockerStatus to console
Write-Host "$(Get-BitLockerStatus | Format-Table | Out-String)" -ForegroundColor 'Yellow' -BackgroundColor 'Black'

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
