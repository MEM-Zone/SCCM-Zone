<#
.SYNOPSIS
    Gets registry key values.
.DESCRIPTION
    Gets registry key values for use in a CI as remediation script.
.NOTES
    Created by
        Ioan Popovici   2018-10-23
#>

[PSObject]$RegistryKeys = @(
    @{
        Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\mrxsmb10'
        Name = 'Start'
    }
    @{
        Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation'
        Name = 'DependOnService'
    }
    @{
        Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters'
        Name = 'SMB1'
    }
    @{
        Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
        Name = 'LmCompatibilityLevel'
    }
)

[PSObject]$Result = @()

$RegistryKeys | ForEach-Object {
    [String]$Value = Get-ItemProperty -Path $_.Path -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty $_.Name -ErrorAction 'SilentlyContinue'

    If (-not $Value) { $Value = 'N/A' }

    [HashTable]$ResultProps = @{
        #Path = $_.Path
        Name = $_.Name
        Value = -Join (' = ', $Value)
    }
    $Result += New-Object 'PSObject' -Property $ResultProps
}

[String]$Output = $($Result | Format-Table -Property Name, Value -HideTableHeaders | Out-String) -replace ('\s+\=\s+', ' = ')

Write-Output -InputObject $Output