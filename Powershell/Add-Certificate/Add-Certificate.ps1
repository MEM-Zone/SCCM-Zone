<#
.SYNOPSIS
    Adds a certificate to the certificate store.
.DESCRIPTION
    Adds a certificate to the certificate store using the certificate key in base64 format.
.EXAMPLE
    Add-Certificate.ps1
.INPUTS
    None.
.OUTPUTS
    System.String.
.NOTES
    Created by Ioan Popovici
.LINK
    https://SCCM.Zone/Add-Certificate-CREDIT (FTW)
.LINK
    https://SCCM.Zone/Add-Certificate
.LINK
    https://SCCM.Zone/Add-Certificate-CHANGELOG
.LINK
    https://SCCM.Zone/Add-Certificate-GIT
.LINK
    https://SCCM.Zone/Issues
.COMPONENT
    Certificate Store
.FUNCTIONALITY
    Add certificate
#>

## Set script requirements
#Requires -Version 3.0

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

## Result object initialization
[psObject]$Result = @()
## Certificate variables
[array]$cerStores =@('Root','TrustedPublisher')
[string]$cerStringBase64 = '
    MIIC7TCCAdWgAwIBAgIQYexQKvQO66dOug2InrN2ZzANBgkqhkiG9w0BAQsFADAm
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    R1TFx1baj97rlziBt2XVZYG9tEFpPxRPD4A5FjRCix/Q
'

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Add-Certificate
Function Add-Certificate {
<#
.SYNOPSIS
    This function is used to add a certificate to the certificate store.
.DESCRIPTION
    This function is used to add a certificate to the certificate store using the certificate base64 key.
.PARAMETER cerStringBase64
    The certificate key to add in base64 format.
.PARAMETER cerStoreLocation
    The certificate Store Location to search. The Default value used is 'LocalMachine'.
    Available Values:
        CurrentUser
        LocalMachine
.PARAMETER cerStoreName
    The certificate Store Name to search.
    Available Values for CurentUser:
        ACRS
        SmartCardRoot
        Root
        Trust
        AuthRoot
        CA
        UserDS
        Disallowed
        My
        TrustedPeople
        TrustedPublisher
        ClientAuthIssuer
    Available Values for LocalMachine:
        rustedPublisher
        ClientAuthIssuer
        Remote Desktop
        Root
        TrustedDevices
        WebHosting
        CA
        WSUS
        Request
        AuthRoot
        TrustedPeople
        My
        SmartCardRoot
        Trust
        Disallowed
        SMS
.EXAMPLE
    Add-Certificate -cerStringBase64 $cerStringBase64 -cerStoreName 'TrustedPublisher'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://SCCM.Zone
.LINK
    https://SCCM.Zone/Git
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('cString')]
        [string]$cerStringBase64,
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('cLocation')]
        [string]$cerStoreLocation = 'LocalMachine',
        [Parameter(Mandatory=$true,Position=2)]
        [Alias('cStore')]
        [string]$cerStoreName
    )

    ## Create certificate store object
    $cerStore = New-Object System.Security.Cryptography.X509Certificates.X509Store $cerStoreName, $cerStoreLocation -ErrorAction 'Stop'

    ## Open the certificate store as Read/Write
    $cerStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)

    ## Convert the base64 string
    $certByteArray = [System.Convert]::FromBase64String($cerStringBase64)

    ## Create new certificate object
    $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ErrorAction 'Stop'

    ## Add certificate to the store
    $Certificate.Import($certByteArray)
    $cerStore.Add($Certificate)

    ## Close the certificate store
    $cerStore.Close()
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

## Cycle specified certificate stores and add the specified certificate
ForEach ($cerStore in $cerStores) {
    Try {
        Add-Certificate -cerStringBase64 $cerStringBase64 -cerStoreName $cerStore -ErrorAction 'Stop'

        #  Create the Result Props
        $ResultProps = [ordered]@{
            'Store' = $cerStore
            'Status'  = 'Add Certificate - Success!'
        }

        #  Adding ResultProps hash table to result object
        $Result += New-Object 'PSObject' -Property $ResultProps
    }
    Catch {

        #  Create the Result Props
        $ResultProps = [ordered]@{
            'Store' = $cerStore
            'Status'  = 'Add Certificate - Failed!'
            'Error' = $_
        }

        #  Adding ResultProps hash table to result object
        $Result += New-Object 'PSObject' -Property $ResultProps
    }
}

## Error handling. If we don't write a stdError when the script fails SCCM will return 'Compliant' because the
## Discovery script does not run again after the Remediation script
If ($Error.Count -ne 0) {

    #  Return result object as an error removing table header for cleaner reporting
    $host.ui.WriteErrorLine($($Result | Format-Table -HideTableHeaders | Out-String))
}
Else {

    #  Return result object removing table header for cleaner reporting
    Write-Output -InputObject $($Result | Format-Table -HideTableHeaders | Out-String)
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
