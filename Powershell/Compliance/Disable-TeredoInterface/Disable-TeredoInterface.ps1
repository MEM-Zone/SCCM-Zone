<#
.SYNOPSIS
    Disables the teredo interface.
.DESCRIPTION
    Disables the teredo interface for use in a CI as remediation script.
.NOTES
    Created by
        Ioan Popovici   2018-09-11
.LINK
    https://SCCM.Zone
.LINK
    https://SCCM.Zone/Git
#>

Try {
    $TeredoState = (Get-NetTeredoConfiguration -ErrorAction 'Stop').Type

    If ($TeredoState -ne 'Disabled') {
        Set-NetTeredoConfiguration -Type 'Disabled' -ErrorAction 'Stop'
        $TeredoState = 'Success: Disable Teredo Tunneling Interface.'
    }
}
Catch {
    $TeredoState = "Error: $($_.Exception.Message)"

}
Finally {
    Write-Output -InputObject $TeredoState
}