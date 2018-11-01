<#
.SYNOPSIS
    Runs a task sequence on a local and remote computers.
.DESCRIPTION
    Runs a task sequence on a local and remote computers by using the 'ExecuteProgram' WMI method.
.EXAMPLE
    Invoke-CMClientTaskSequence.ps1 -Computer 'SomeComputerName' -Name 'SomeTaskSequenceName'
.INPUTS
    System.String
.OUTPUTS
    System.String.
.NOTES
    Created by Ioan Popovici
.LINK
    Blog    : https://SCCM-Zone.com
.LINK
    Changes : https://SCCM.Zone/Invoke-CMTaskSequence-CHANGELOG
.LINK
    Github  : https://SCCM.Zone/Invoke-CMTaskSequence
.LINK
    Issues  : https://SCCM.Zone/issues
.COMPONENT
    CM Client
.FUNCTIONALITY
    Run task sequence
#>

## Set script requirements
#Requires -Version 3.0

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

## Get script parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true,Position=0)]
    [Alias('Name')]
    [String]$TaskSequenceName,
    [Parameter(Mandatory=$false,Position=1)]
    [Alias('Computer')]
    [String]$ComputerName = $null
)

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Invoke-CMClientTaskSequence
Function Invoke-CMClientTaskSequence {
<#
.SYNOPSIS
    Runs a task sequence on a local and remote computers.
.DESCRIPTION
    Runs a task sequence on a local and remote computers by using the 'ExecuteProgram' WMI method.
.EXAMPLE
    Invoke-CMClientTaskSequence.ps1 -Computer 'SomeComputerName' -Name 'SomeTaskSequenceName'
.INPUTS
    System.String
.OUTPUTS
    System.String.
.NOTES
    Created by Ioan Popovici
.LINK
    Blog    : https://SCCM-Zone.com
.LINK
    Changes : https://SCCM.Zone/Invoke-CMTaskSequence-CHANGELOG
.LINK
    Github  : https://SCCM.Zone/Invoke-CMTaskSequence
.LINK
    Issues  : https://SCCM.Zone/issues
.COMPONENT
    CM Client
.FUNCTIONALITY
    Run task sequence
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Name')]
        [string]$TaskSequenceName,
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('Computer')]
        [string]$ComputerName = $null
    )

    Begin {
        Try {

            [cimclass]$CimClass = (Get-CimClass -ComputerName $ComputerName -Namespace 'Root\ccm\clientsdk' -ClassName 'CCM_ProgramsManager' -ErrorAction 'Stop')
            [pscustomobject]$TaskSequence = (Get-CimInstance -ComputerName $ComputerName -Namespace 'Root\ccm\clientSDK' -ClassName 'CCM_Program' -ErrorAction 'Stop' | Where-Object -Property 'Name' -like $TaskSequenceName)

            [hashtable]$Arguments = @{
                'PackageID' = $TaskSequence.PackageID
                'ProgramID' = $TaskSequence.ProgramID
            }
            If ((-not $Arguments.PackageID) -or (-not $TaskSequence.ProgramID)) {
                $ErrorMessage = "Could not resolve task sequence [$TaskSequenceName] PackageID or ProgramID."
            }
        }
        Catch {
            $ErrorMessage = "Could not resolve task sequence [$TaskSequenceName]. `n $_.ErrorMesage"
        }
    }
    Process {
        If (-not $Error) {
            Try {
                Invoke-CimMethod -ComputerName $ComputerName -CimClass $CimClass -MethodName 'ExecuteProgram' –Arguments $Arguments -ErrorAction 'Stop'
            }
            Catch {
                $ErrorMessage = "Could not run task sequence [$TaskSequenceName]. `n $_.ErrorMesage"
            }
        }
    }
    End {
        If (-not $Error) {
            Write-Output "Task sequence [$TaskSequenceName] has run successfully"
        }
        Else {
            Write-Error -Message $ErrorMessage
        }
    }
}

#endregion
##*=============================================
##* END FUNCTION LISTINGS
##*=============================================

##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody

Invoke-CMClientTaskSequence -ComputerName $ComputerName -TaskSequenceName $TaskSequenceName

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
