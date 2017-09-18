<#
*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 3.0, AD CommandLets                                  *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2017-09-06 | v1.0     | First version                                                 *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This PowerShell Script is used to change the current password.
.DESCRIPTION
    This PowerShell Script is used to change the current password, circumventing the password history.
.EXAMPLE
    C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe -NoExit -NoProfile -File Set-PreviousPassword.ps1
.NOTES
    The script will change the password 23 times and then set the desired password.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody

## Read variables from console
[string]$Identity = Read-Host -Prompt 'Provide Account'
[string]$OldPassword = Read-Host -Prompt 'Provide Current Password'
[string]$NewPassword = Read-Host -Prompt 'Provide New Password'

## Change password 23 times
For ($i=1; $i -lt 23; $i++) {
    [string]$TmpPassword = $OldPassword + $i
    Try {

        Set-ADAccountPassword -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $TmpPassword -Force) -ErrorAction 'Stop'
        Write-Host "Setting Temporary Password - Success. `n Current Password: $TmpPassword" -ForegroundColor 'Yellow' -BackgroundColor 'Black'
        $OldPassword = $TmpPassword
    }
    Catch {
        Write-Host "Setting Temporary Password - Failed! `n Current Password: $OldPassword" -ForegroundColor 'Red' -BackgroundColor 'Black'
    }
}

## Set new password
Try {
    Write-Host 'Setting Permanent Password...' -ForegroundColor 'Green' -BackgroundColor 'Black'
    Set-ADAccountPassword -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)
    Write-Host "Setting Permanent Password - Success. `n Current Password: $NewPassword" -ForegroundColor 'Green' -BackgroundColor 'Black'
}
Catch {
    Write-Host "Setting Permanent Password - Failed! `n Current Password: $OldPassword" -ForegroundColor 'Red' -BackgroundColor 'Black'
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
