<#
.SYNOPSIS
    This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
    The script is provided as a template to perform an install or uninstall of an application(s).
    The script either performs an "Install" deployment type or an "Uninstall" deployment type.
    The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
    The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
    The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
    Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
    Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
    Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
    Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
    Toolkit Exit Code Ranges:
    60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
    69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
    70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
    http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false)]
    [ValidateSet('Install','Uninstall')]
    [string]$DeploymentType = 'Install',
    [Parameter(Mandatory=$false)]
    [ValidateSet('Interactive','Silent','NonInteractive')]
    [string]$DeployMode = 'Interactive',
    [Parameter(Mandatory=$false)]
    [switch]$AllowRebootPassThru = $false,
    [Parameter(Mandatory=$false)]
    [switch]$TerminalServerMode = $false,
    [Parameter(Mandatory=$false)]
    [switch]$DisableLogging = $false
)

Try {
    ## Set the script execution policy for this process
    Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}

    ##*===============================================
    ##* VARIABLE DECLARATION
    ##*===============================================
    ## Variables: Application
    [string]$appVendor = 'Microsoft Corporation'
    [string]$appName = 'SCCM Client'
    [string]$appVersion = '5.00.8458.1500'
    [string]$appArch = 'ALL'
    [string]$appLang = 'EN'
    [string]$appRevision = '01.00'
    [string]$appScriptVersion = '1.0.0'
    [string]$appScriptDate = '2017-02-10'
    [string]$appScriptAuthor = 'Ioan Popovici'
    ##*===============================================
    ## Variables: Install Titles (Only set here to override defaults set by the toolkit)
    [string]$installName = ''
    [string]$installTitle = ''

    ##* Do not modify section below
    #region DoNotModify

    ## Variables: Exit Code
    [int32]$mainExitCode = 0

    ## Variables: Script
    [string]$deployAppScriptFriendlyName = 'Deploy Application'
    [version]$deployAppScriptVersion = [version]'3.6.9'
    [string]$deployAppScriptDate = '02/12/2017'
    [hashtable]$deployAppScriptParameters = $psBoundParameters

    ## Variables: Environment
    If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
        [string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

    ## Dot source the required App Deploy Toolkit Functions
    Try {
        [string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
        If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
        If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
    }
    Catch {
        If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
        Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
        ## Exit the script, returning the exit code to SCCM
        If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
    }

    #endregion
    ##* Do not modify section above
    ##*===============================================
    ##* END VARIABLE DECLARATION
    ##*===============================================

        If ($deploymentType -ine 'Uninstall') {
        ##*===============================================
        ##* PRE-INSTALLATION
        ##*===============================================
        [string]$installPhase = 'Pre-Installation'

        ## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
        Show-InstallationWelcome -CloseApps 'ccmsetup=SCCM Client Setup' -CheckDiskSpace -PersistPrompt

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Installation tasks here>
        #  Get Certificate from the cache
        $CertificateFileName = Get-Item -Path $DirFiles\Certificates\*.pfx  | Select-Object -First 1 -ExpandProperty Name

        #  Get the SubjectName from the Certificate File Name excluding everything after the first dot
        $CertificateFileHostName = ($CertificateFileName | Select-String -Pattern '[^.]*').Matches.Value

        #  Check if Certificate is in the cache
        If ($CertificateFileName -eq $Null) {

            #  If Certificate is not found ask user to generate it
            $UserResponse = Show-DialogBox -Title "Generate Certificate" -Text "Client Certificate not found! Generate Certificate?`n`nWARNING:`n`nMUST BE RUN ON A DATAKRAFVERK DOMAIN/SUBDOMAIN MACHINE ADDED TO SCCM!" -Buttons "YesNo"  -DefaultButton "Second" -Icon "Exclamation"

            #  Check user response and ask user for Certificate SubjectName and Export Password
            If($UserResponse -eq "Yes") {

                #  Get Certificate SubjectName and make sure that user input is not null
                Do {
                    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                    $CertificateHostname = [Microsoft.VisualBasic.Interaction]::InputBox('Please enter Hostname for which to generate Certificate:','Enter NetBiosName/FQDN.')
                }
                While ($CertificateHostname.Length -eq 0)

                #  Get Certificate Export Password and make sure that user input is not null. Notify user that UserName is not Needed
                Do {
                    $CertificatePassword =  $Host.ui.PromptForCredential('Certificate Password', 'Please enter a Password for the Client Certificate.', 'UserName Not Needed!', 'NetBiosUserName')
                }
                While ($CertificatePassword.Length -eq 0)

                #  Make Certificate Request and store the generated Certificate in the Personal Certificate Store
                $Certificate = Get-Certificate -Template ConfigMgrClientWindowsCertificateforExport -SubjectName CN=$CertificateHostName -CertStoreLocation Cert:\LocalMachine\My

                #  Get Certificate Thumbprint
                $CertificateThumbprint = $Certificate.Certificate.Thumbprint


                #  Export Certificate to the Cache using Thumbprint for lookup and previous entered Password
                Export-PfxCertificate –Cert Cert:\LocalMachine\My\$CertificateThumbprint –FilePath $DirFiles\Certificates\$CertificateHostName.pfx -ChainOption EndEntityCertOnly -NoProperties -Password $CertificatePassword.Password

                #  Remove Certificate from Personal Store
                Remove-Item -Path Cert:\LocalMachine\My\$CertificateThumbprint

                #  Notify user that Certificate is generated and added to cache.
                Show-DialogBox -Title "SCCM Client Certificate" -Text "Client Certificate generated and added to cache! `n`n`nCopy this application to $CertificateHostname and run it again." -Icon "Information"

                #  Exit Script Gracefully
                Exit-Script -ExitCode 0

            #  Check if user response is 'No' and Exit Script Gracefully
            } Else { Exit-Script -ExitCode 0 }

        }

        #  Check if SubjectName matches the current Hostname
        ElseIf ($Env:ComputerName -ne $CertificateFileHostName) {

            #  Notify the user that Certificate SubjectName does not match the current Hostname
            Show-DialogBox -Title "Certificate Subject Name Mismatch!" -Text "Certificate Name does not match Machine Hostname!`n`n`nCopy this application to $CertificateFileHostName and run it again." -Icon "Exclamation"

            #  Ask user if he wants to clear the Certificate cache
            $UserResponse = Show-DialogBox -Title "Delete Certificate" -Text "Delete any existing Certificate from Cache?" -Buttons "YesNo" -DefaultButton "Second" -Icon "Question"

            #  Clear Certificate cache if user response is 'Yes'
            If($UserResponse -eq "Yes") {
                Remove-Item -Path $DirFiles\*.pfx -Force

                #  Exit Script Gracefully
                Exit-Script -ExitCode 0

            #  If user response is 'No' Exit Script Gracefully
            } Else { Exit-Script -ExitCode 0 }

        }

        #  If there are no Certificate Prerequisites are met, proceed with Certificate Import
        Else {

            #  Get Certificate Export Password and make sure that user input is not null. Notify user that UserName is not Needed
            Do {
                $CertificatePassword =  $Host.ui.PromptForCredential('Certificate Password', 'Please enter the SCCM Client Certificate Password.', 'UserName Not Needed!', 'NetBiosUserName')
            }
            While ($CertificatePassword.Length -eq 0)

            #  Import Root CA, Intermediary CA and Client Certificates, specify to continue on error if CA Certificates fail to import as a dirty fix when we don't
            #  have rights to overwrite exiting CA Certificates. Must be addressed in the future
            Import-Certificate -FilePath "$DirFiles\Certificates\VismaRootCA.cer" -CertStoreLocation Cert:\LocalMachine\Root -ErrorAction 'Continue'
            Import-Certificate -FilePath "$DirFiles\Certificates\VismaIntCA.cer" -CertStoreLocation Cert:\LocalMachine\Root -ErrorAction 'Continue'
            Import-PfxCertificate –FilePath "$DirFiles\Certificates\$CertificateFileName" -CertStoreLocation Cert:\LocalMachine\My -Password $CertificatePassword.Password
        }

        ##*===============================================
        ##* INSTALLATION
        ##*===============================================
        [string]$installPhase = 'Installation'

        ## Handle Zero-Config MSI Installations
        If ($useDefaultMsi) {
        [hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
        Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
        }

        ## <Perform Installation tasks here>
        Show-BalloonTip -BalloonTipIcon "Info" -BalloonTipText "Installing SCCM Client..."

        #  Installing SCCM Client
        Execute-Process -FilePath "CCMSETUP.EXE" -Parameters "/Source:$DirFiles /NoService /UsePKICert /NoCRLCheck CCMHOSTNAME=VITCSCCMCMG.CLOUDAPP.NET/CCM_Proxy_MutualAuth/72057594037927955 SMSSITECODE=VIT CCMALWAYSINF=1 CCMFIRSTCERT=1 SMSSIGNCERT=$DirFiles\Certificates\SMS_Signing_Certificate.cer "

        #  Write OS Bitness to log for troubleshooting purposes
        Write-Log $is64Bit

        ##*===============================================
        ##* POST-INSTALLATION
        ##*===============================================
        [string]$installPhase = 'Post-Installation'

        ## <Perform Post-Installation tasks here>
        #  Prompt user fir SCEP Installation with a timeout of 30 seconds. After this time 'No' will be automatically selected
        $UserResponse = Show-DialogBox -Title "Install SCEP" -Text "Do you want to install SCEP?" -Buttons "YesNo"  -DefaultButton "First" -Icon "Question" -Timeout "30"

        #  Check user response and proceed with the installation if response is 'Yes'
        If($UserResponse -eq "Yes") {
            Show-BalloonTip -BalloonTipIcon "Info" -BalloonTipText "Installing SCEP..."

            #  Set the SCEP Policy Configuration File path
            $SCEPPolicyPath = "$dirFiles"+"\ep_defaultpolicy.xml"

            #  Start SCEP installation
            Execute-Process -FilePath "scepinstall.exe " -Arguments "/policy $SCEPPolicyPath /s /q" -IgnoreExitCodes "-2147156218" -WindowStyle 'Hidden'
        }

        ## Display a message at the end of the install
        If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'Installation was Successful!' -ButtonRightText 'OK' -Icon Information -NoWait }
    }
    ElseIf ($deploymentType -ieq 'Uninstall')
    {
        ##*===============================================
        ##* PRE-UNINSTALLATION
        ##*===============================================
        [string]$installPhase = 'Pre-Uninstallation'

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        Show-InstallationWelcome -CloseApps 'ccmsetup=SCCM Client Setup' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Uninstallation tasks here>


        ##*===============================================
        ##* UNINSTALLATION
        ##*===============================================
        [string]$installPhase = 'Uninstallation'

        ## Handle Zero-Config MSI Uninstallations
        If ($useDefaultMsi) {
            [hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
            Execute-MSI @ExecuteDefaultMSISplat
        }

        ## <Perform Uninstallation tasks here>
        #  Uninstall SCEP
        Show-BalloonTip -BalloonTipText 'Uninstalling SCEP...' -BalloonTipTitle 'Uninstalling SCEP'
        Write-Log 'Uninstalling SCEP...'
        Execute-Process -Path 'scepinstall.exe' -Parameters '/u /s' -WindowStyle 'Hidden'

        #  Uninstall SCCM Client
        Show-BalloonTip -BalloonTipText 'SCCM Client...' -BalloonTipTitle 'Uninstalling CCM'
        Write-Log 'Uninstalling SCCM...'
        Execute-Process -Path 'CCMSETUP.EXE' -Parameters '/uninstall' -WindowStyle 'Hidden'

        ##*===============================================
        ##* POST-UNINSTALLATION
        ##*===============================================
        [string]$installPhase = 'Post-Uninstallation'

        ## <Perform Post-Uninstallation tasks here>


    }

    ##*===============================================
    ##* END SCRIPT BODY
    ##*===============================================

    ## Call the Exit-Script function to perform final cleanup operations
    Exit-Script -ExitCode $mainExitCode
}
Catch {
    [int32]$mainExitCode = 60001
    [string]$mainErrorMessage = "$(Resolve-Error)"
    Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
    Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
    Exit-Script -ExitCode $mainExitCode
}
