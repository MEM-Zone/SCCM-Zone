<#
**********************************************************************************************************
*                                                                                                        *
*** This Powershell Script is used to get AD computer information from a list of forests  			   ***
*                                                                                                        *
**********************************************************************************************************
* Created by Ioan Popovici, 13/11/2015  | Requirements Powershell 3.0   								 *
* =======================================================================================================*
* Modified by   |    Date    | Revision |                            Comments                            *
*________________________________________________________________________________________________________*
* Ioan Popovici | 13/11/2015 | v1.0      | First version                                 				 *
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
If ((Test-Path $ResultPath) -eq $False) {
	New-Item -Path $ResultPath -Type Directory | Out-Null
  Out-File -Path $ResultPath+"GetADForestComputers.log"
} ElseIf (Test-Path $ResultPath) {
		Remove-Item $ResultPath\* -Recurse -Force
    $ErrorLog = $ResultPath+"GetADForestComputers.log"
    $ErrorLog | Out-File -Path -Force
	}

## Check if Forrest csv exists
$ADForrestListCSV = $PSScriptRoot+"ADForrestList.csv"
If ((Test-Path $ADForrestListCSV) -eq $False) {
  "ADForrest.csv does not exist!" | Out-File -Path $ResultPath+"GetADForestComputers.log"
  Write-Host "ADForrest.csv does not exist!"
  Exit
} Else {
    $ADForrestList = Import-Csv -Path $ADForrestListCSV
  }

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

## Process Imported CSV Object Data
$ADForrestList | ForEach-Object {

	#  Initialize variables
  $ADForest = $null
	$ADForestDomains = $null
	$Domain = $null

  #  Get AD Forrest domains
	Try {
		$ADForest = Get-ADForest $Forest -ErrorAction SilentlyContinue -ErrorVariable Error
		$ADForestDomains = $ADForest.Domains
	}
	Catch {
	   Write-Log -Message "Failed to connect to forest: $Forest, $Error"
	}
	If ($ADForestDomains -ne $null) {
		ForEach ($Domain in $ADForestDomains) {
			Try {
				Get-ADComputer -Server $Domain -Filter {Enabled -eq $true} -Property * -ErrorVariable Error | Select-Object Name, OperatingSystem, @{Name='Domain';Expression={$Domain}} | Export-Csv "$ResultPath\ADForestComputers.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
			}
			Catch {
				Write-Log -Message "No permissions to domain: $Domain, $Error"
			}
		}
	}
}
Write-Host ""
Write-Log -Message "Processing Finished!" -BackgroundColor Green -ForegroundColor White

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
