<#
**********************************************************************************************************
*                                                                                                        *
*** This Powershell Script is used to get AD computer information from a list of forests               ***
*                                                                                                        *
**********************************************************************************************************
* Created by Ioan Popovici, 13/11/2015  | Requirements Powershell 3.0                                    *
* =======================================================================================================*
* Modified by   |    Date    | Revision | Comments                                                       *
*________________________________________________________________________________________________________*
* Ioan Popovici | 13/11/2015 | v1.0     | First version                                                  *
* Ioan Popovici | 22/02/2016 | v1.1     | Improved Logging                                               *
* Ioan Popovici | 22/02/2016 | v1.2     | Added Progress Indicator                                       *
*--------------------------------------------------------------------------------------------------------*
*                                                                                                        *
**********************************************************************************************************

    .SYNOPSIS
        This Powershell Script is used to get AD computer information from a list of forests.
    .DESCRIPTION
        This Powershell Script is used to get AD computer name, operating system and domain from a list of forests.
#>

##*=============================================
##* INITIALIZATION
##*=============================================
#region Initialization

## Initialize Logging
$ResultPath = $PSScriptRoot+"\Results"
$ErrorLog = $PSScriptRoot+"\GetADForestComputers.log"

#  Create Result Directory
If ((Test-Path $ResultPath) -eq $False) {
    New-Item -Path $ResultPath -Type Directory | Out-Null
} ElseIf (Test-Path $ResultPath) {

        #  Clean Result Directory
        Remove-Item $ResultPath\* -Recurse -Force
    }

#  Clean Log
Get-Date | Out-File $ErrorLog -Force

#endregion
##*=============================================
##* END INITIALIZATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Write-Log
Function Write-Log {
<#
.SYNOPSIS
    Write messages to a log file in CMTrace.exe compatible format or Legacy text file format.
.DESCRIPTION
    Write messages to a log file in CMTrace.exe compatible format or Legacy text file format and optionally display in the console.
.PARAMETER Message
    The message to write to the log file or output to the console.
.EXAMPLE
    Write-Log -Message "Error"
.NOTES
.LINK
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Text')]
        [string]$Message
    )

    ## Getting the Date and time
    $DateAndTime = Get-Date

    ### Writing to log file
    "$DateAndTime : $Message" | Out-File $ErrorLog -Append
    "$DateAndTime : $_" | Out-File $ErrorLog -Append

    ## Writing to Console
    Write-Host $Message  -ForegroundColor Red -BackgroundColor White
    Write-Host $_.Exception -ForegroundColor White -BackgroundColor Red
    Write-Host ""
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

## Clearing Console
CLS

## Initialize forest counter
$ProgressCounterForest = 0

## Check if ADForestList.csv exists and import the content into $ADForestList variable
$ADForestListCSV = $PSScriptRoot+"\ADForestList.csv"
If ((Test-Path $ADForestListCSV) -eq $False) {
  Write-Log -Message "ADForestList.csv does not exist!"
  Write-Host "ADForest.csv does not exist!"

  #  Exit script if ADForestList.csv is not found
  Exit
} Else {

    #  Import CSV content into $ADForestList variable
    $ADForestList = (Import-Csv -Path $ADForestListCSV).Forest
  }

## Process Imported CSV Forest List
ForEach ($Forest in $ADForestList) {

    #  Initialize variables
      $ADForest = $null
    $ADForestDomains = $null
    $Domain = $null
    $ProgressCounterDomain = 0

    #  Show Forest progress bar
    If ($($ADForestList.Count) -ne $null) {
        $ProgressCounterForest++
        $PercentCompleteForest = (($ProgressCounterForest / $ADForestList.Count) * 100)
        Write-Progress -Id 1 -Activity "Processing Forest: $Forest" -Status "$ProgressCounterForest of $($ADForestList.Count) Forests" -CurrentOperation "$PercentCompleteForest% complete" -PercentComplete $PercentCompleteForest
    }

      #  Get AD Forest domains
    Try {
        $ADForest = Get-ADForest $Forest -ErrorAction SilentlyContinue -ErrorVariable Error1
        $ADForestDomains = $ADForest.Domains
    }
    Catch {
       Write-Log -Message "Failed to connect to forest: $Forest, $ErrorVar"
    }

    ## Process Forest domains with error handling
    If ($ADForestDomains -ne $null) {
        ForEach ($Domain in $ADForestDomains) {
            Try {

                #  Show Domain progress bar
                If ($($ADForestDomains.Count) -ne $null) {
                    $ProgressCounterDomain++
                    $PercentCompleteDomain = (($ProgressCounterDomain / $ADForestDomains.Count) * 100)
                    Write-Progress -Id 2 -Activity "Processing Domain: $Domain" -Status "$ProgressCounterDomain of $($ADForestDomains.Count) Forest Domains" -CurrentOperation "$PercentCompleteDomain% complete" -PercentComplete $PercentCompleteDomain
                }

                ## Get domain computers
                $ADComputers = Get-ADComputer -Server $Domain -Filter {Enabled -eq $true} -Property * -ErrorVariable $ErrorVar | Select-Object Name, OperatingSystem, @{Name='Domain';Expression={$Domain}}

                ## Reset computer progress bar
                $ProgressCounterComputers = 0

                ## Export computers to CSV file
                ForEach ($Computer in $ADComputers) {

                    #  Show Computer progress bar
                    If ($($ADComputers.Count) -ne $null) {
                        $ProgressCounterComputers++
                        $PercentCompleteComputer = '{0:N0}' -f (($ProgressCounterComputers / $($ADComputers.Count)) * 100)
                        Write-Progress -Id 3 -Activity "Processing Computer: $($Computer.Name)" -Status "$ProgressCounterComputers of $($ADComputers.Count) Domain Computers" -CurrentOperation "$PercentCompleteComputer% complete" -PercentComplete $PercentCompleteComputer
                    }

                    #  Write Computer to CSV file
                    $Computer | Export-Csv "$ResultPath\ADForestComputers.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
                }
            }
            Catch {
                Write-Log -Message "No permissions to domain: $Domain, $ErrorVar"
            }

        }
    }
}
Write-Host ""
Write-Log -Message "Processing Finished!"

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
