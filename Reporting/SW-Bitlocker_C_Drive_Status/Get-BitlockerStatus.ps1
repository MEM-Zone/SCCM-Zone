<#
*********************************************************************************************************
*                                                                                                       *
*** This Powershell Script is used to get the Bitlocker  status                                       ***
*                                                                                                       *
*********************************************************************************************************
* Created by Ioan Popovici, 13/11/2015  | Requirements Powershell 3.0                                   *
* ======================================================================================================*
* Modified by   |    Date    | Revision | Comments                                                      *
*_______________________________________________________________________________________________________*
* Ioan Popovici | 13/11/2015 | v1.0     | First version                                                 *
*-------------------------------------------------------------------------------------------------------*
*                                                                                                       *
*********************************************************************************************************

    .SYNOPSIS
        This Powershell Script is used to get the Bitlocker protection status.
    .DESCRIPTION
        This Powershell Script is used to get the Bitlocker protection status for C drive.
#>

##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody

## Get the Bitlocker Encryption Status for C drive
Try {

  #  Read the status from wmi
  Get-WmiObject -Namespace “root\CIMV2\Security\MicrosoftVolumeEncryption” -Class Win32_EncryptableVolume -ErrorAction Stop | `
    ForEach-Object {
      $ID = $_.DriveLetter;

      #  Make it more report friendly
      Switch ($_.GetProtectionStatus().ProtectionStatus) {
        0 { $State = "PROTECTION OFF" } 1 { $State = "PROTECTION ON"} 2 { $State = "PROTECTION UNKNOWN"}
      }

      #  Check if protection is on for C drive
      If (($ID -eq “C:”) -and ($State -eq "PROTECTION ON")) {
        $Protection = $true
      }
    }
}

## Catch any script errors
Catch {
  $ScriptError = $true
}

## Write protection status to console
If ($Protection) {
  Write-Host "PROTECTION ON"
} ElseIf ($ScriptError -ne $true) {
    Write-Host "PROTECTION OFF"
  } Else {
      Write-Host "SCRIPT EXECUTION ERROR!"
    }

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
