<#
*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 3.0, AD CommandLets                                  *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2017-09-06 | v1.0     | First version                                                 *
* Ioan Popovici | 2017-11-06 | v1.1     | Added Random Password Generator                               *
* Ioan Popovici | 2018-05-14 | v1.2     | Secured passwords                                             *
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

# Initialize Progress Counter
$ProgressCounter = 0

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Show-InputPrompt
Function Show-InputPrompt {
<#
.SYNOPSIS
    Displays a custom input prompt with optional buttons.
.DESCRIPTION
    Any combination of Left, Middle or Right buttons can be displayed. The return value of the button clicked by the user is the button text specified.
.PARAMETER Title
    Title of the prompt.
.PARAMETER Label
    Label text to be included in the prompt.
.PARAMETER LabelAlignment
    Alignment of the label text. Options: Left, Center, Right. Default: Left.
.PARAMETER Text
    Text to be included in the Text box prompt
.PARAMETER TextAlignment
    Alignment of the Text Box text. Options: Left, Center, Right. Default: Left.
.PARAMETER ButtonLeftText
    Show a button on the left of the prompt with the specified text.
.PARAMETER ButtonRightText
    Show a button on the right of the prompt with the specified text.
.PARAMETER ButtonMiddleText
    Show a button in the middle of the prompt with the specified text.
.PARAMETER MinimizeWindows
    Specifies whether to minimize other windows when displaying prompt. Default: $false.
.EXAMPLE
    Show-InputPrompt -Title 'Domains or Domain Controllers' -Label 'Input Domains or Domain Controllers:' -Text 'Domains go here.' -ButtonRightText 'Ok' -ButtonLeftText 'Cancel'
.NOTES
    Function modified from original source
.LINK
    http://psappdeploytoolkit.com
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        [ValidateNotNullorEmpty()]
        [string]$Title = '',
        [Parameter(Mandatory=$false)]
        [string]$Label = '',
        [Parameter(Mandatory=$false)]
        [ValidateSet('Left','Center','Right')]
        [string]$LabelAlignment = 'Left',
        [Parameter(Mandatory=$false)]
        [string]$Text = '',
        [Parameter(Mandatory=$false)]
        [ValidateSet('Left','Center','Right')]
        [string]$TextAlignment = 'Left',
        [Parameter(Mandatory=$false)]
        [string]$ButtonRightText = '',
        [Parameter(Mandatory=$false)]
        [string]$ButtonLeftText = '',
        [Parameter(Mandatory=$false)]
        [string]$ButtonMiddleText = '',
        [Parameter(Mandatory=$false)]
        [ValidateNotNullorEmpty()]
        [boolean]$MinimizeWindows = $false
    )

    [Windows.Forms.Application]::EnableVisualStyles()
    $formInputPrompt = New-Object -TypeName 'System.Windows.Forms.Form'
    $labeltextDomain = New-Object -TypeName 'System.Windows.Forms.Label'
    $textBoxDomain = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $labeltextIdentity = New-Object -TypeName 'System.Windows.Forms.Label'
    $textBoxIdentity = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $labeltextPassword = New-Object -TypeName 'System.Windows.Forms.Label'
    $textBoxPassword = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $labeltextNewPassword = New-Object -TypeName 'System.Windows.Forms.Label'
    $textBoxNewPassword = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $labeltextIterations = New-Object -TypeName 'System.Windows.Forms.Label'
    $textBoxIterations = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $labeltextSleep = New-Object -TypeName 'System.Windows.Forms.Label'
    $textBoxSleep = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $buttonRight = New-Object -TypeName 'System.Windows.Forms.Button'
    $buttonMiddle = New-Object -TypeName 'System.Windows.Forms.Button'
    $buttonLeft = New-Object -TypeName 'System.Windows.Forms.Button'
    $buttonAbort = New-Object -TypeName 'System.Windows.Forms.Button'
    $InitialformInputPromptWindowState = New-Object -TypeName 'System.Windows.Forms.FormWindowState'

    [scriptblock]$Form_Cleanup_FormClosed = {
        ## Remove all event handlers from the controls
        Try {
            $textBox.remove_Click($handler_textBox_Click)
            $buttonLeft.remove_Click($buttonLeft_OnClick)
            $buttonRight.remove_Click($buttonRight_OnClick)
            $buttonMiddle.remove_Click($buttonMiddle_OnClick)
            $buttonAbort.remove_Click($buttonAbort_OnClick)
            $formInputPrompt.remove_Load($Form_StateCorrection_Load)
            $formInputPrompt.remove_FormClosed($Form_Cleanup_FormClosed)
        }
        Catch { }
    }

    [scriptblock]$Form_StateCorrection_Load = {
        ## Correct the initial state of the form to prevent the .NET maximized form issue
        $formInputPrompt.WindowState = 'Normal'
        $formInputPrompt.AutoSize = $true
        $formInputPrompt.TopMost = $true
        $formInputPrompt.BringToFront()
        # Get the start position of the form so we can return the form to this position if PersistPrompt is enabled
        Set-Variable -Name 'formInputPromptStartPosition' -Value $formInputPrompt.Location -Scope 'Script'
    }

    ## Form
    ##----------------------------------------------
    ## Create padding object
    $paddingNone = New-Object -TypeName 'System.Windows.Forms.Padding'
    $paddingNone.Top = 0
    $paddingNone.Bottom = 0
    $paddingNone.Left = 0
    $paddingNone.Right = 0

    ## Generic Label properties
    $labelPadding = '20,0,20,0'

    ## Generic Text properties
    $textPadding = '20,0,20,0'

    ## Generic Object Spacing
    $ObjectSpacing = 20
    $ObjectSpacingIncrement = 25

    ## Generic Button properties
    $buttonWidth = 110
    $buttonHeight = 23
    $buttonPadding = 50
    $buttonSize = New-Object -TypeName 'System.Drawing.Size'
    $buttonSize.Width = $buttonWidth
    $buttonSize.Height = $buttonHeight
    $buttonPadding = New-Object -TypeName 'System.Windows.Forms.Padding'
    $buttonPadding.Top = 0
    $buttonPadding.Bottom = 5
    $buttonPadding.Left = 50
    $buttonPadding.Right = 0

    ## Domain
    #  Domain Label Text
    $labeltextDomain.DataBindings.DefaultDataSourceUpdateMode = 0
    $labeltextDomain.Name = 'labeltextDomain'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 455
    $labeltextDomain.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 4
    $System_Drawing_Point.Y = $ObjectSpacing
    $labeltextDomain.Location = $System_Drawing_Point
    $labeltextDomain.Margin = '0,0,0,0'
    $labeltextDomain.Padding = $labelPadding
    $labeltextDomain.TabIndex = 1
    $labeltextDomain.Text = 'Domain or Server'
    $labeltextDomain.TextAlign = "Middle$($LabelAlignment)"
    $labeltextDomain.Anchor = 'Top'
    $labeltextDomain.add_Click($handler_labeltextDomain_Click)

    #  Domain Text Box
    $textBoxDomain.DataBindings.DefaultDataSourceUpdateMode = 0
    $textBoxDomain.Name = 'textBoxDomain'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 330
    $System_Drawing_Size.Width = 390
    $textBoxDomain.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $textBoxDomain.Location = $System_Drawing_Point
    $textBoxDomain.Margin = '0,0,0,0'
    $textBoxDomain.Padding = $textPadding
    $textBoxDomain.TabIndex = 2
    $textBoxDomain.Text = $env:USERDNSDOMAIN
    $textBoxDomain.TextAlign = $TextAlignment
    $textBoxDomain.Anchor = 'Top'
    $textBoxDomain.add_Click($handler_textBoxDomain_Click)

    ## Identity
    #  Identity Label Text
    $labeltextIdentity.DataBindings.DefaultDataSourceUpdateMode = 0
    $labeltextIdentity.Name = 'labeltextIdentity'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 455
    $labeltextIdentity.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 4
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $labeltextIdentity.Location = $System_Drawing_Point
    $labeltextIdentity.Margin = '0,0,0,0'
    $labeltextIdentity.Padding = $labelPadding
    $labeltextIdentity.TabIndex = 1
    $labeltextIdentity.Text = 'Identity'
    $labeltextIdentity.TextAlign = "Middle$($LabelAlignment)"
    $labeltextIdentity.Anchor = 'Top'
    $labeltextIdentity.add_Click($handler_labeltextIdentity_Click)

    #  Identity Text Box
    $textBoxIdentity.DataBindings.DefaultDataSourceUpdateMode = 0
    $textBoxIdentity.Name = 'textBoxIdentity'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 330
    $System_Drawing_Size.Width = 390
    $textBoxIdentity.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $textBoxIdentity.Location = $System_Drawing_Point
    $textBoxIdentity.Margin = '0,0,0,0'
    $textBoxIdentity.Padding = $textPadding
    $textBoxIdentity.TabIndex = 2
    $textBoxIdentity.Text = $env:USERNAME
    $textBoxIdentity.TextAlign = $TextAlignment
    $textBoxIdentity.Anchor = 'Top'
    $textBoxIdentity.add_Click($handler_textBoxIdentity_Click)

    #  Password Label Text
    $labeltextPassword.DataBindings.DefaultDataSourceUpdateMode = 0
    $labeltextPassword.Name = 'labeltextPassword'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 455
    $labeltextPassword.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 4
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $labeltextPassword.Location = $System_Drawing_Point
    $labeltextPassword.Margin = '0,0,0,0'
    $labeltextPassword.Padding = $labelPadding
    $labeltextPassword.TabIndex = 1
    $labeltextPassword.Text = 'Password'
    $labeltextPassword.TextAlign = "Middle$($LabelAlignment)"
    $labeltextPassword.Anchor = 'Top'
    $labeltextPassword.add_Click($handler_labeltextPassword_Click)

    #  Password Text Box
    $textBoxPassword.DataBindings.DefaultDataSourceUpdateMode = 0
    $textBoxPassword.Name = 'textBoxPassword'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 330
    $System_Drawing_Size.Width = 390
    $textBoxPassword.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $textBoxPassword.Location = $System_Drawing_Point
    $textBoxPassword.Margin = '0,0,0,0'
    $textBoxPassword.Padding = $textPadding
    $textBoxPassword.TabIndex = 2
    $textBoxPassword.Text = $null
    $textBoxPassword.PasswordChar = '*'
    $textBoxPassword.TextAlign = $TextAlignment
    $textBoxPassword.Anchor = 'Top'
    $textBoxPassword.add_Click($handler_textBoxPassword_Click)

    ## New Password
    #  New Password Label Text
    $labeltextNewPassword.DataBindings.DefaultDataSourceUpdateMode = 0
    $labeltextNewPassword.Name = 'labeltextNewPassword'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 455
    $labeltextNewPassword.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 4
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $labeltextNewPassword.Location = $System_Drawing_Point
    $labeltextNewPassword.Margin = '0,0,0,0'
    $labeltextNewPassword.Padding = $labelPadding
    $labeltextNewPassword.TabIndex = 1
    $labeltextNewPassword.Text = 'New Password'
    $labeltextNewPassword.TextAlign = "Middle$($LabelAlignment)"
    $labeltextNewPassword.Anchor = 'Top'
    $labeltextNewPassword.add_Click($handler_labeltextNewPassword_Click)

    #  New Password Text Box
    $textBoxNewPassword.DataBindings.DefaultDataSourceUpdateMode = 0
    $textBoxNewPassword.Name = 'textBoxNewPassword'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 330
    $System_Drawing_Size.Width = 390
    $textBoxNewPassword.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $textBoxNewPassword.Location = $System_Drawing_Point
    $textBoxNewPassword.Margin = '0,0,0,0'
    $textBoxNewPassword.Padding = $textPadding
    $textBoxNewPassword.TabIndex = 2
    $textBoxNewPassword.Text = $null
    $textBoxNewPassword.PasswordChar = '*'
    $textBoxNewPassword.TextAlign = $TextAlignment
    $textBoxNewPassword.Anchor = 'Top'
    $textBoxNewPassword.add_Click($handler_textBoxNewPassword_Click)


    ## Iterations
    #  Iterations Label Text
    $labeltextIterations.DataBindings.DefaultDataSourceUpdateMode = 0
    $labeltextIterations.Name = 'labeltextIterations'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 455
    $labeltextIterations.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 4
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $labeltextIterations.Location = $System_Drawing_Point
    $labeltextIterations.Margin = '0,0,0,0'
    $labeltextIterations.Padding = $labelPadding
    $labeltextIterations.TabIndex = 1
    $labeltextIterations.Text = 'Number of Iterations to perform'
    $labeltextIterations.TextAlign = "Middle$($LabelAlignment)"
    $labeltextIterations.Anchor = 'Top'
    $labeltextIterations.add_Click($handler_labeltextIterations_Click)

    #  Iterations Text Box
    $textBoxIterations.DataBindings.DefaultDataSourceUpdateMode = 0
    $textBoxIterations.Name = 'textBoxIterations'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 330
    $System_Drawing_Size.Width = 390
    $textBoxIterations.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $textBoxIterations.Location = $System_Drawing_Point
    $textBoxIterations.Margin = '0,0,0,0'
    $textBoxIterations.Padding = $textPadding
    $textBoxIterations.TabIndex = 2
    $textBoxIterations.Text = '30'
    $textBoxIterations.TextAlign = $TextAlignment
    $textBoxIterations.Anchor = 'Top'
    $textBoxIterations.add_Click($handler_textBoxIterations_Click)

    ## Sleep Time
    #  Sleep Time Label Text
    $labeltextSleep.DataBindings.DefaultDataSourceUpdateMode = 0
    $labeltextSleep.Name = 'labeltextSleep'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 455
    $labeltextSleep.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 4
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $labeltextSleep.Location = $System_Drawing_Point
    $labeltextSleep.Margin = '0,0,0,0'
    $labeltextSleep.Padding = $labelPadding
    $labeltextSleep.TabIndex = 1
    $labeltextSleep.Text = 'Sleep Time in Seconds between Iterations'
    $labeltextSleep.TextAlign = "Middle$($LabelAlignment)"
    $labeltextSleep.Anchor = 'Top'
    $labeltextSleep.add_Click($handler_labeltextSleep_Click)

    #  Sleep Time Text Box
    $textBoxSleep.DataBindings.DefaultDataSourceUpdateMode = 0
    $textBoxSleep.Name = 'textBoxSleep'
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 330
    $System_Drawing_Size.Width = 390
    $textBoxSleep.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object -TypeName 'System.Drawing.Point'
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = $ObjectSpacing = $ObjectSpacing + $ObjectSpacingIncrement
    $textBoxSleep.Location = $System_Drawing_Point
    $textBoxSleep.Margin = '0,0,0,0'
    $textBoxSleep.Padding = $textPadding
    $textBoxSleep.TabIndex = 2
    $textBoxSleep.Text = '10'
    $textBoxSleep.TextAlign = $TextAlignment
    $textBoxSleep.Anchor = 'Top'
    $textBoxSleep.add_Click($handler_textBoxSleep_Click)

    ## Button Left
    $buttonLeft.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonLeft.Location = '15,400'
    $buttonLeft.Name = 'buttonLeft'
    $buttonLeft.Size = $buttonSize
    $buttonLeft.TabIndex = 3
    $buttonLeft.Text = $buttonLeftText
    $buttonLeft.DialogResult = 'No'
    $buttonLeft.AutoSize = $false
    $buttonLeft.UseVisualStyleBackColor = $true
    $buttonLeft.add_Click($buttonLeft_OnClick)

    ## Button Middle
    $buttonMiddle.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonMiddle.Location = '170,400'
    $buttonMiddle.Name = 'buttonMiddle'
    $buttonMiddle.Size = $buttonSize
    $buttonMiddle.TabIndex = 4
    $buttonMiddle.Text = $buttonMiddleText
    $buttonMiddle.DialogResult = 'Ignore'
    $buttonMiddle.AutoSize = $true
    $buttonMiddle.UseVisualStyleBackColor = $true
    $buttonMiddle.add_Click($buttonMiddle_OnClick)

    ## Button Right
    $buttonRight.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonRight.Location = '325,400'
    $buttonRight.Name = 'buttonRight'
    $buttonRight.Size = $buttonSize
    $buttonRight.TabIndex = 5
    $buttonRight.Text = $ButtonRightText
    $buttonRight.DialogResult = 'Yes'
    $buttonRight.AutoSize = $true
    $buttonRight.UseVisualStyleBackColor = $true
    $buttonRight.add_Click($buttonRight_OnClick)

    ## Button Abort (Hidden)
    $buttonAbort.DataBindings.DefaultDataSourceUpdateMode = 0
    $buttonAbort.Name = 'buttonAbort'
    $buttonAbort.Size = '1,1'
    $buttonAbort.DialogResult = 'Abort'
    $buttonAbort.TabStop = $false
    $buttonAbort.UseVisualStyleBackColor = $true
    $buttonAbort.add_Click($buttonAbort_OnClick)

    ## Form Input Prompt
    $System_Drawing_Size = New-Object -TypeName 'System.Drawing.Size'
    $System_Drawing_Size.Height = 400
    $System_Drawing_Size.Width = 455
    $formInputPrompt.Size = $System_Drawing_Size
    $formInputPrompt.Padding = '0,0,0,10'
    $formInputPrompt.Margin = $paddingNone
    $formInputPrompt.DataBindings.DefaultDataSourceUpdateMode = 0
    $formInputPrompt.Name = 'WelcomeForm'
    $formInputPrompt.Text = $title
    $formInputPrompt.StartPosition = 'CenterScreen'
    $formInputPrompt.FormBorderStyle = 'FixedDialog'
    $formInputPrompt.MaximizeBox = $false
    $formInputPrompt.MinimizeBox = $false
    $formInputPrompt.TopMost = $true
    $formInputPrompt.TopLevel = $true
    #  Domain
    $formInputPrompt.Controls.Add($labeltextDomain)
    $formInputPrompt.Controls.Add($textBoxDomain)
    #  Identity
    $formInputPrompt.Controls.Add($labeltextIdentity)
    $formInputPrompt.Controls.Add($textBoxIdentity)
    #  Password
    $formInputPrompt.Controls.Add($labeltextPassword)
    $formInputPrompt.Controls.Add($textBoxPassWord)
    #  NewPassword
    $formInputPrompt.Controls.Add($labeltextNewPassword)
    $formInputPrompt.Controls.Add($textBoxNewPassword)
    #  Iterations
    $formInputPrompt.Controls.Add($labeltextIterations)
    $formInputPrompt.Controls.Add($textBoxIterations)
    #  Sleep
    $formInputPrompt.Controls.Add($labeltextSleep)
    $formInputPrompt.Controls.Add($textBoxSleep)
    #  Buttons
    $formInputPrompt.Controls.Add($buttonAbort)
    If ($buttonLeftText) { $formInputPrompt.Controls.Add($buttonLeft) }
    If ($buttonMiddleText) { $formInputPrompt.Controls.Add($buttonMiddle) }
    If ($buttonRightText) { $formInputPrompt.Controls.Add($buttonRight) }

    ## Save the initial state of the form
    $InitialformInputPromptWindowState = $formInputPrompt.WindowState
    ## Init the OnLoad event to correct the initial state of the form
    $formInputPrompt.add_Load($Form_StateCorrection_Load)
    ## Clean up the control events
    $formInputPrompt.add_FormClosed($Form_Cleanup_FormClosed)
    ## Show the prompt synchronously. If user cancels, then keep showing it until user responds using one of the buttons and enters some text.
    $showDialog = $true
    While ($showDialog) {
        # Minimize all other windows
        If ($minimizeWindows) { $null = $shellApp.MinimizeAll() }
        # Show the Form
        $result = $formInputPrompt.ShowDialog()
        # Validate txtboxes
        $textBoxesEmpty  =
        If ($textBoxDomain.Text -and
            $textBoxIdentity.Text -and
            $textBoxPassword.Text -and
            $textBoxNewPassword.Text -and
            $textBoxIterations.Text -and
            $textBoxSleep.Text
        ) { $false } Else { $true }

        If (($result -eq 'Yes' -and $textBoxesEmpty -eq $false) -or ($result -eq 'No') -or ($result -eq 'Ignore') -or ($result -eq 'Abort')) {
            $showDialog = $false
        }
    }

    $formInputPrompt.Dispose()

    ##  Create output object
    [PSCustomObject]$Output = [ordered]@{
        'Domain' = $textBoxDomain.Text
        'Identity' = $textBoxIdentity.Text
        'Password' = $textBoxPassword.Text
        'NewPassword' = $textBoxPassword.Text
        'Iterations' = $textBoxIterations.Text
        'Sleep' = $textBoxSleep.Text
        'Button' = $result
    }

    Write-Output -InputObject $Output
}
#endregion

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

## Get input data
$FormData = Show-InputPrompt -Title 'Change Password' -ButtonRightText 'Ok' -ButtonLeftText 'Cancel'

## Exit if 'Cancel' is pressed
If ($FormData.Button -eq 'NO') { Break }

## Set variables
[string]$Domain = $FormData.Domain
[string]$Identity = $FormData.Identity
[string]$OldPassword = $FormData.Password
[string]$NewPassword = $FormData.NewPassword
[string]$Iterations = $FormData.Iterations
[string]$Sleep = $FormData.Sleep
#  Set credentials
[securestring]$OldPasswordSecure = ConvertTo-SecureString $OldPassword -AsPlainText -Force
[System.Management.Automation.PSCredential]$Credential = New-Object System.Management.Automation.PSCredential ($Identity, $OldPasswordSecure)

## Change password
Write-Host 'Setting temporary passwords...' -ForegroundColor 'Green' -BackgroundColor 'Black'

For ($i=1; $i -le $Iterations; $i++) {
    [string]$RandomPassword = Get-RandomPassword
    Try {
        #  Write Progress
        $ProgressCounter++
        Write-Progress -Activity 'Change Password' -CurrentOperation "Changing Password [$RandomPassword]" -PercentComplete (($ProgressCounter / $Iterations) * 100)

        #  Change Password
        Set-ADAccountPassword -Server $Domain -Identity $Identity -OldPassword (ConvertTo-SecureString -ASPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $RandomPassword -Force) -ErrorAction 'Stop'
        Write-Host "[$i/$Iterations] Current password is: $RandomPassword" -ForegroundColor 'Yellow' -BackgroundColor 'Black'
        $OldPassword = $RandomPassword
        Start-Sleep -Seconds $Sleep
    }
    Catch {
        Write-Error -Message "Failed to set temporary password. `n$_" -Category 'SecurityError' -ErrorAction 'Stop'
    }
}

## Set new password
Try {
    Write-Host 'Setting permanent password...' -ForegroundColor 'Green' -BackgroundColor 'Black'
    Set-ADAccountPassword -Server $Domain -Identity $Identity -OldPassword (ConvertTo-SecureString -AsPlainText $OldPassword -Force) -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)
    Write-Host "Current password set succesfully" -ForegroundColor 'Green' -BackgroundColor 'Black'
}
Catch {
    Write-Error -Message "Failed to set temporary password. `n$_" -Category 'SecurityError' -ErrorAction 'Stop'
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================