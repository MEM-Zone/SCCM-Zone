<#
.SYNOPSIS
    Gets the BitLocker protection status.
.DESCRIPTION
    Gets the BitLocker protection status for a specific drive, or all drives.
.PARAMETER DriveType
    Specifies the drive type(s) for which to get the bitlocker status. Default is: '3'.
    Available values
        0   DRIVE_UNKNOWN
        1   DRIVE_NO_ROOT_DIR
        2   DRIVE_REMOVABLE
        3   DRIVE_FIXED
        4   DRIVE_REMOTE
        5   DRIVE_CDROM
        6   DRIVE_RAMDISK
    These values are just for reference you probably will never use them.
.PARAMETER DriveLetter
    Specifies the drive letter(s) for which to get the bitlocker status. Default is: 'All'.
.PARAMETER ShowTableHeaders
    This switch specifies to show the table headers. Default: $false.
.EXAMPLE
    Get-BitLockerStatus.ps1 -DriveLetter 'All'
.EXAMPLE
    Get-BitLockerStatus.ps1 -DriveType '2'
.EXAMPLE
    Get-BitLockerStatus.ps1 -DriveType '2','3' -DriveLetter 'C:','D:'
.INPUTS
    System.String.
.OUTPUTS
    System.String.
.NOTES
    Created by
        Ioan Popovici   2015-11-15
    Release notes
        https://github.com/Ioan-Popovici/SCCMZone/blob/master/Powershell/Get-BitlockerStatus/CHANGELOG.md
    For issue reporting please use github
        https://github.com/Ioan-Popovici/SCCMZone/issues
.LINK
    https://SCCM-Zone.com
.LINK
    https://github.com/Ioan-Popovici/SCCMZone
.COMPONENT
    BitLocker
.FUNCTIONALITY
    Get BitLocker status
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

## Get script parameters
Param (
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateNotNullorEmpty()]
    [Alias('Type')]
    [string[]]$DriveType = '3',
    [Parameter(Mandatory = $false, Position = 1)]
    [ValidateNotNullorEmpty()]
    [Alias('Drive')]
    [string[]]$DriveLetter = 'All',
    [Parameter(Mandatory = $false, Position = 2)]
    [ValidateNotNullorEmpty()]
    [Alias('ShowHeaders')]
    [switch]$ShowTableHeaders = $false
)

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

## Set table headers
[boolean]$HideTableHeaders = If ($ShowTableHeaders) { $false } Else { $true }

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
.PARAMETER DriveType
    Specifies the drive type(s) for which to get the bitlocker status. Default is: '3'.
    Available values
        0   DRIVE_UNKNOWN
        1   DRIVE_NO_ROOT_DIR
        2   DRIVE_REMOVABLE
        3   DRIVE_FIXED
        4   DRIVE_REMOTE
        5   DRIVE_CDROM
        6   DRIVE_RAMDISK
    These values are just for reference you probably will never use them.
.PARAMETER DriveLetter
    Specifies the drive letter(s) for which to get the bitlocker status. Default is: 'All'.
.EXAMPLE
    Get-BitLockerStatus -DriveLetter 'All'
.EXAMPLE
    Get-BitLockerStatus -DriveType '2'
.EXAMPLE
    Get-BitLockerStatus -DriveType '2','3' -DriveLetter 'C:','D:'
.INPUTS
    System.String.
.OUTPUTS
    System.Object.
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://SCCM-Zone.com
.LINK
    https://github.com/Ioan-Popovici/SCCMZone
.COMPONENT
    BitLocker
.FUNCTIONALITY
    Get BitLocker status
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateNotNullorEmpty()]
        [Alias('Type')]
        [string[]]$DriveType = '3',
        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullorEmpty()]
        [Alias('Drive')]
        [string[]]$DriveLetter = 'All'
    )

    Begin {

        ## Initializing Result Object
        [psCustomObject]$Result = @()
    }
    Process {
        Try {

            ##  Get the local drives from WMI
            [psObject]$LocalDrives = Get-CimInstance -Namespace 'root\CIMV2' -ClassName 'CIM_LogicalDisk' | Where-Object -Property 'DriveType' -in $DriveType

            ## Get the BitLocker Status for all drives from WMI
            Get-CimInstance  -Namespace 'root\CIMV2\Security\MicrosoftVolumeEncryption'  -ClassName 'Win32_EncryptableVolume' -ErrorAction 'Stop' | `
                Where-Object -Property 'DriveLetter' -in $($LocalDrives.DeviceID) | `
                ForEach-Object {

                #  Get the drive type
                [string]$GetDriveType = $($LocalDrives | Where-Object -Property 'DeviceID' -eq $($_.DriveLetter)) | Select-Object -ExpandProperty 'DriveType'

                #  Create the Result Props and make the ProtectionStatus more report friendly
                [hashtable]$ResultProps = [ordered]@{
                    'Drive'            = $($_.DriveLetter)
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
                    'DriveType' = 'Type ' + $GetDriveType
                }

                #  Adding ResultProps hash table to result object
                $Result += New-Object PSObject -Property $ResultProps
            }

            #  Workaround for some Windows 7 computers not reporting BitLocker protection status for all drives
            #  Create the ResultProps array
            $LocalDrives | ForEach-Object {
                If ($($_.DeviceID) -notin $($Result.Drive)) {
                    $ResultProps = [ordered]@{
                        'Drive'            = $($_.DeviceID)
                        'ProtectionStatus' = 'PROTECTION OFF'
                        'DriveType'        = $($_.DriveType)
                    }

                    #  Adding ResultProps hash table to result object
                    $Result += New-Object PSObject -Property $ResultProps
                }
            }
        }

        ## Catch any script errors
        Catch {
            Write-Error -Message "Script Execution Error!`n $_" -Category 'NotSpecified'
        }
        Finally {

            ## Filter result depending the DriveLetter parameter
            If ($DriveLetter -ne 'All') {
                Write-Output -InputObject $($Result | Where-Object -Property 'Drive' -in $DriveLetter)
            }
            Else {
                Write-Output -InputObject $Result
            }
        }
    }
    End {
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
[string]$Result = $(Get-BitLockerStatus -DriveType $DriveType -DriveLetter $DriveLetter | Format-Table -HideTableHeaders:$HideTableHeaders | Out-String)
Write-Output $Result

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================