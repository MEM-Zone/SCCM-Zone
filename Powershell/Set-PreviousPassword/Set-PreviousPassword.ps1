<#
*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 3.0, AD CommandLets                                  *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2017-09-06 | v1.0     | First version                                                 *
* Ioan Popovici | 2017-11-06 | v1.1     | Added Random Password Generator                               *
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
    The script will change the password 31 times and then set the desired password.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

## Read variables from console
[string]$Identity = Read-Host -Prompt 'Provide Account'
[string]$OldPassword = Read-Host -Prompt 'Provide Current Password'
[string]$NewPassword = Read-Host -Prompt 'Provide New Password'

#endregion
##*=============================================##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Get-RandomPassword
Function Get-RandomPassword() {
<#.SYNOPSIS
    Generates a random password.
.DESCRIPTION
    Generates a random strong password.
.PARAMETER passLength
    The generated password length.
.PARAMETER passSource
    The character source used for passowrd generation.
.EXAMPLE
    Get-RandomPassword -passLength '20'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    Credit to:
    https://blogs.technet.microsoft.com/heyscriptingguy/2013/06/03/generating-a-new-password-with-windows-powershell/
.LINK
    https://sccm-zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCM
#>
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [Alias('pLength')]
        [int]$passLength=30,
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('pSource')]
        [string[]]$passSource = $(
            $ascii=$NULL
            For ($a=33; $a –le 126; $a++) { $ascii +=, [char][byte]$a }
            Write-Output $ascii
        )
    )

    ## Generate random password using password source
    For ($loop = 1; $loop –le $passLength; $loop++) {
        $Result += ($passSource | Get-Random)
    }

    ## Return result to pipeline
    Write-Output $Result
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

## Change password 30 times
Write-Host 'Setting temporary passwords...' -ForegroundColor 'Green' -BackgroundColor 'Black'

1..31 | ForEach {
    [string]$RandomPassword = Get-RandomPassword
    Try {
        Set-ADAccountPassword -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $RandomPassword -Force) -ErrorAction 'Stop'
        Write-Host "Current password is: $RandomPassword" -ForegroundColor 'Yellow' -BackgroundColor 'Black'
        $OldPassword = $RandomPassword
        Start-Sleep -Seconds 1
    }
    Catch {
        Write-Host "Failed to set temporary password. `n Current password is: $OldPassword" -ForegroundColor 'Red' -BackgroundColor 'Black'
        Break
    }
}

## Set new password
Try {
    Write-Host 'Setting permanent password...' -ForegroundColor 'Green' -BackgroundColor 'Black'
    Set-ADAccountPassword -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)
    Write-Host "Current password is: $NewPassword" -ForegroundColor 'Green' -BackgroundColor 'Black'
}
Catch {
    Write-Host "Failed to set permanent password. `n$_ `n Current password is: $OldPassword" -ForegroundColor 'Red' -BackgroundColor 'Black'
    Break
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================