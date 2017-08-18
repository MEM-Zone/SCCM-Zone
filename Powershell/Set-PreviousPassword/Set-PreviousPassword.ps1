$Identity = Read-Host -Prompt 'Provide Account'
$OldPassword = Read-Host -Prompt 'Provide Current Password'
$NewPassword = Read-Host -Prompt 'Provide New Password'
For ($i=1; $i -lt 23; $i++) {
   [string]$TmpPassword = $OldPassword + $i
   $OldPassword

   Set-ADAccountPassword -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $oldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $TmpPassword -Force)
   $OldPassword = $TmpPassword
}

Set-ADAccountPassword -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)
