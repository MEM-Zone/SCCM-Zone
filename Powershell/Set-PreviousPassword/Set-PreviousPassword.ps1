$Identity = Read-Host -Prompt 'Provide Account'
$OldPassword = Read-Host -Prompt 'Provide Current Password'
$NewPassword = Read-Host -Prompt 'Provide New Password'
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

Try {
    Write-Host 'Setting Permanent Password...' -ForegroundColor 'Green' -BackgroundColor 'Black'
    Set-ADAccountPassword -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)
    Write-Host "Setting Permanent Password - Success. `n Current Password: $NewPassword" -ForegroundColor 'Green' -BackgroundColor 'Black'
}
Catch {
    Write-Host "Setting Permanent Password - Failed! `n Current Password: $OldPassword" -ForegroundColor 'Red' -BackgroundColor 'Black'
}
