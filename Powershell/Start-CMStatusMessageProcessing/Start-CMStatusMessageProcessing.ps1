<#
**********************************************************************************************************
*                                                                                                        *
*** This Powershell Script is used to set send mails when SCCM security scopes change  				   ***
*                                                                                                        *
**********************************************************************************************************
* Created by Ioan Popovici, 30/03/2015  | Requirements Powershell 4.0   								 *
* =======================================================================================================*
* Modified by   |    Date    | Revision |                            Comments                            *
*________________________________________________________________________________________________________*
* Ioan Popovici | 29/09/2015 | v1.0      | First version                                 				 *
*--------------------------------------------------------------------------------------------------------*
*                                                                                                        *
**********************************************************************************************************

	.SYNOPSIS
        This Powershell Script is used to set send mails when SCCM security scopes change.
    .DESCRIPTION
        This Powershell Script is used to set send mails when SCCM security Production and Quality Security Scopes change.    
#>


# SCCM Status message parameter
Param (
	[string]$SMDescription
)

# Send email function
Function Send-Mail {
	Param (
		[Parameter(Mandatory=$false)]
		[string]$From = "SCCM Site Server <noreply@visma.com>", 
		[Parameter(Mandatory=$false)]
		[string]$To = "SCCM Team <SCCM-Team@visma.com>",
		[Parameter(Mandatory=$false)]
		[string]$CC = "",
		[Parameter(Mandatory=$false)]
		[string]$Subject = "Info: Quality Check Needed!",
		[Parameter(Mandatory=$true)]
		[string]$Body,
		[Parameter(Mandatory=$false)]
		[string]$SMTPServer = "mail.datakraftverk.no",
		[Parameter(Mandatory=$false)]
		[string]$SMTPPort = "25"
	)	
	Try {
		Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -ErrorAction 'Stop'	
	}
	Catch {
		Write-Error "Send Mail Failed!"
	}
}

	# Arrays for Name, RegEx, and result object
	$NameArray = @('User','Object','Type','Scope','Action')
	$PatternArray = @('\S+(?<=\\)\S+','(?<=object\s)[^(]+','(?<=Type:\s..._)\w+','(?<=scope:\s)\w+','(associated|deassociated)')
	$InfoArray = @{}

	# RegEx pattern matching
	ForEach ($Item in 0..($NameArray.length -1)){
		$SMDescription | Select-String -Pattern $PatternArray[$Item] -AllMatches | % { $InfoArray.($NameArray[$Item]) = $_.Matches.Value }
	}

	# Building object from result array 
	$Result = New-Object -TypeName PSObject -Property $InfoArray
	Write-Output $Result

	# Send diferent mails depending on Scope and action
	If (($Result.Scope -eq "Quality") -and ($Result.Action -eq "associated")) {
		Send-Mail -Body "$($Result.User) $($Result.Action) $($Result.Scope) Scope to the $($Result.Type) named $($Result.Object)" 
	} 
	ElseIf ($Result.Scope -eq "Production" -and ($Result.Action -eq "associated")) {
		Send-Mail -Subject "Warning: Production Scope Added!" -Body "$($Result.User) $($Result.Action) $($Result.Scope) Scope to the $($Result.Type) named $($Result.Object)"
	}
	Else {
		Write-Host "No actions match, nothing to do..."
	}