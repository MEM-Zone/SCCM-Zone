<#
.SYNOPSIS
    Selects a certificate in the certificate store.
.DESCRIPTION
    Selects a certificate in a specified certificate store using the certificate Serial Number.
.EXAMPLE
    Select-Certificate.ps1
.NOTES
    Created by Ioan Popovici
.LINK
    Credit  : https://SCCM.Zone/Add-Certificate-CREDIT (FTW)
.LINK
    BlogPost: https://SCCM.Zone/Add-Certificate
.LINK
    Changes : https://SCCM.Zone/Select-Certificate-CHANGELOG
.LINK
    Github  : https://SCCM.Zone/Select-Certificate-GIT
.LINK
    Issues  : https://SCCM.Zone/Issues
.COMPONENT
    Certificate Store
.FUNCTIONALITY
    Select certificate
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
[string]$cerSerialNumber = '6xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx7'
#  Remove spaces from certificate serial number
$cerSerialNumber = $cerSerialNumber -replace '\s',''

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Select-Certificate
Function Select-Certificate {
<#
.SYNOPSIS
    This function is used to get the details of a specific certificate.
.DESCRIPTION
    This function is used to get the details of a Specific certificate using the certificate 'Serial Number'.
.PARAMETER cerSerialNumber
    The certificate Serial Number to search for.
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
    Select-Certificate -cerSerialNumber '61ec50244f40eeba74eba0d889eb37667' -cerStoreName 'TrustedPublisher'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    Github  : https://SCCM.Zone/Select-Certificate-GIT
.LINK
    Issues  : https://SCCM.Zone/Issues
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('cSerial')]
        [string]$cerSerialNumber,
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('cLocation')]
        [string]$cerStoreLocation = 'LocalMachine',
        [Parameter(Mandatory=$true,Position=2)]
        [Alias('cStore')]
        [string]$cerStoreName
    )

    ## Create certificate store object
    $cerStore = New-Object System.Security.Cryptography.X509Certificates.X509Store $cerStoreName, $cerStoreLocation -ErrorAction 'Stop'

    ## Open the certificate store as ReadOnly
    $cerStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

    ## Get the certificate details
    $Result = $cerStore.Certificates | Where-Object { $_.SerialNumber -eq $cerSerialNumber }  | Select-Object SerialNumber,Thumbprint,Subject,Issuer,NotBefore,NotAfter

    ## Close the certificate Store
    $cerStore.Close()

    ## Return certificate details or a 'Certificate Selection - Failed!' string if the certificate does not exist
    If ($Result) {
        Write-Output -InputObject $Result
    }
    Else {
        Write-Output -InputObject 'Certificate Selection - Failed!'
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

## Cycle specified certificate stores and search for specified certificate serial number
$cerStores | ForEach-Object {
    $GetCertificate = Select-Certificate -cerSerialNumber $cerSerialNumber -cerStoreName $_

    #  Create the Result Props and make the GetCertificate more report friendly
    $ResultProps = [ordered]@{
        'Store' = $_
        'Certificate' = $(
            Switch ($GetCertificate) {
                {$GetCertificate.SerialNumber} {"$($GetCertificate.SerialNumber) - Found!"}
                Default { $GetCertificate }
            }
        )
    }

    #  Adding ResultProps hash table to result object
    $Result += New-Object 'PSObject' -Property $ResultProps
}

## Workaround for SCCM Compliance Rule limitation. The remediation checkbox shows up only if 'Equals' rule is specified.
If ($($Result.Certificate | Out-String) -notmatch 'Failed') {

    #  Return 'Compliant'
    Write-Output -InputObject 'Compliant'
}
Else {

#  Return result object removing table header for cleaner reporting
Write-Output -InputObject $($Result | Format-Table -HideTableHeaders | Out-String)
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
