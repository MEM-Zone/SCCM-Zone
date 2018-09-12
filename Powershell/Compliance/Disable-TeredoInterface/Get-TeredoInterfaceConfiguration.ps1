<#
.SYNOPSIS
    This PowerShell script is used get the teredo tunneling interface configuration state.
.DESCRIPTION
    This PowerShell script is used get the teredo tunneling interface configuration state for use in a CI as detection script.
.NOTES
    Created by:
        Ioan Popovici   2018-09-11
#>

Try {
    $TeredoState = (Get-NetTeredoConfiguration -ErrorAction 'Stop').Type
}
Catch {
    $TeredoState = "Error: $($_.Exception.Message)"

}
Finally {
    Write-Output -InputObject $TeredoState
}