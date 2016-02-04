<#
*********************************************************************************************************
*                                                                                                       *
*** This Powershell Script is used to clean the CCM cache of all non persisted content                ***
*                                                                                                       *
*********************************************************************************************************
* Created by Ioan Popovici, 13/11/2015  | Requirements Powershell 3.0                                   *
* ======================================================================================================*
* Modified by   |    Date    | Revision |                            Comments                           *
*_______________________________________________________________________________________________________*
* Ioan Popovici | 13/11/2015 | v1.0     | First version                                                 *
* Ioan Popovici | 16/11/2015 | v1.1     | Improved Logging                                              *
* Ioan Popovici | 16/11/2015 | v1.2     | Vastly Improved                                               *
* Ioan Popovici | 03/02/2016 | v2.0     | Vastly Improved                                               *
* Ioan Popovici | 04/02/2016 | v2.1     | Fixed TotalSize decimals                                      *
*-------------------------------------------------------------------------------------------------------*
*                                                                                                       *
*********************************************************************************************************

	.SYNOPSIS
        This Powershell Script is used to clean the CCM cache of all non persisted content.
    .DESCRIPTION
        This Powershell Script is used to clean the CCM cache of all non persisted content that is not needed anymore.
		It only cleans packages, applications and updates that have a installed status and are not persisted, other
		cache items will NOT be cleaned.
#>

##*=============================================
##* INITIALIZATION
##*=============================================
#region Initialization

## Cleaning prompt history
CLS

## Global variable
$Global:Result  =@()

## Initalize progress Counter
$ProgressCounter = 0

## Configure Logging
# Set log path
$ResultCSV = "C:\Temp\Clean-CCMCache.log"

# Remove previous log it it's more than 500 KB
If (Test-Path $ResultCSV) {
	If ((Get-Item $ResultCSV).Length -gt 500KB) {
		Remove-Item $ResultCSV -Force | Out-Null
	}
}

# Get log parent path
$ResultPath =  Split-Path $ResultCSV -Parent

# Create path directory if it does not exist
If ((Test-Path $ResultPath) -eq $False) {
	New-Item -Path $ResultPath -Type Directory | Out-Null
}

## Get the current date
$Date = Get-Date

## Get list of all non persisted content in CCMCache, only this content will be removed
$CM_CacheItems = Get-WmiObject -Namespace root\ccm\SoftMgmtAgent -Query 'SELECT ContentID,Location FROM CacheInfoEx WHERE PersistInCache = 0'


#endregion
##*=============================================
##* END INITIALIZATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Remove-CacheItem
Function Remove-CacheItem {
<#
.SYNOPSIS
	Removes SCCM cache item if it's not persisted.
.DESCRIPTION
	Removes specified SCCM cache item if it's not found in the persisted cache list.
.PARAMETER CacheItemToDelete
	The cache item ID that needs to be deleted.
.PARAMETER CacheItemName
	The cache item name that needs to be deleted.
.EXAMPLE
	Remove-CacheItem -CacheItemToDelete "{234234234}" -CacheItemName "Office2003"
.NOTES
.LINK
#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		[Alias('CacheTD')]
		[string]$CacheItemToDelete,
		[Parameter(Mandatory=$true,Position=1)]
		[Alias('CacheN')]
		[string]$CacheItemName
	)

	## Delete chache item if it's non persisted
	If ($CM_CacheItems.ContentID -contains $CacheItemToDelete) {

		# Get Cache item location and size
		$CacheItemLocation = $CM_CacheItems | Where {$_.ContentID -Contains $CacheItemToDelete} | Select -ExpandProperty Location
		$CacheItemSize =  Get-ChildItem $CacheItemLocation -Recurse -Force | Measure-Object -Property Length -Sum | Select -ExpandProperty Sum

		# Build result object
		$ResultProps = [ordered]@{
			'DeletionDate'	= $Date
			'Name' = $CacheItemName
			'ID' = $CacheItemToDelete
			'Location' = $CacheItemLocation
			'Size(MB)' = "{0:N2}" -f ($CacheItemSize / 1MB)
			'TotalDeleted(MB)' = ''
		}

		# Add items to result object
		$Global:Result  += New-Object PSObject -Property $ResultProps

		# Connect to resource manager COM object
		$CMObject = New-Object -ComObject "UIResource.UIResourceMgr"

	 	# Use GetCacheInfo method to return cache properties
		$CMCacheObjects = $CMObject.GetCacheInfo()

		# Delete Cache element
		$CMCacheObjects.GetCacheElements() | Where-Object {$_.ContentID -eq $CacheItemToDelete} |
			ForEach-Object {
				$CMCacheObjects.DeleteCacheElement($_.CacheElementID)
			}
	}
}
#endregion

#region Function Remove-CachedApplications
Function Remove-CachedApplications {
<#
.SYNOPSIS
	Removes SCCM cached application.
.DESCRIPTION
	Removes specified SCCM cache application if it's already installed.
.PARAMETER
.EXAMPLE
.NOTES
.LINK
#>

	## Get list of applications
	$CM_Applications = Get-WmiObject -Namespace root\ccm\ClientSDK -Query 'SELECT * FROM CCM_Application'

	## Check for installed applications
	Foreach ($Application in $CM_Applications) {

		## Show progrss bar
		If ($CM_Applications.Count -ne $null) {
			$ProgressCounter++
			Write-Progress -Activity 'Processing Applications' -CurrentOperation $Application.FullName -PercentComplete (($ProgressCounter / $CM_Applications.Count) * 100)
			Start-Sleep -Milliseconds 400
		}
		## Get Application Properties
		$Application.Get()

	    ## Enumerate all deployment types for an application
	    Foreach ($DeploymentType in $Application.AppDTs) {
	        If ($Application.InstallState -eq "Installed" -and $Application.IsMachineTarget) {

				## Get content ID for specific application deployment type
	            $AppType = "Install",$DeploymentType.Id,$DeploymentType.Revision
	            $Content = Invoke-WmiMethod -Namespace root\ccm\cimodels -Class CCM_AppDeliveryType -Name GetContentInfo -ArgumentList $AppType

				## Call Remove-CacheItem function
				Remove-CacheItem -CacheTD $Content.ContentID -CacheN $Application.FullName
	        }
	    }
	}
}
#endregion

#region Function Remove-CachedPackages
Function Remove-CachedPackages {
<#
.SYNOPSIS
	Removes SCCM cached package.
.DESCRIPTION
	Removes specified SCCM cached package if it's not needed anymore.
.PARAMETER
.EXAMPLE
.NOTES
.LINK
#>

	## Get list of packages
	$CM_Packages = Get-WmiObject -Namespace root\ccm\ClientSDK -Query 'SELECT PackageID,PackageName,LastRunStatus,RepeatRunBehavior FROM CCM_Program'

	## Check if any deployed programs in the package need the cached package and add deletion or exemption list for comparison
	ForEach ($Program in $CM_Packages) {

		# Check if program in the package needs the cached package
		If ($Program.LastRunStatus -eq "Succeeded" -and $Program.RepeatRunBehavior -ne "RerunAlways" -and $Program.RepeatRunBehavior -ne "RerunIfSuccess") {

			# Add PackageID to Deletion List if not already added
			If ($Program.PackageID -NotIn $PackageIDDeleteTrue) {
				[array]$PackageIDDeleteTrue += $Program.PackageID
			}

		} Else {

				# Add PackageID to Exception List if not already added
				If ($Program.PackageID -NotIn $PackageIDDeleteFalse) {
				[array]$PackageIDDeleteFalse += $Program.PackageID
			}
		}
	}

	## Parse Deletion List and Remove Package if not in Exemption List
	ForEach ($Package in $PackageIDDeleteTrue) {

		# Show progress bar
		If ($CM_Packages.Count -ne $null) {
			$ProgressCounter++
			Write-Progress -Activity 'Processing Packages' -CurrentOperation $Package.PackageName -PercentComplete (($ProgressCounter / $CM_Packages.Count) * 100)
			Start-Sleep -Milliseconds 800
		}
		# Call Remove Function if Package is not in $PackageIDDeleteFalse
		If ($Package -NotIn $PackageIDDeleteFalse) {
			Remove-CacheItem -CacheTD $Package.PackageID -CacheN $Package.PackageName
		}
	}
}
#endregion

#region Function Remove-CachedUpdates
Function Remove-CachedUpdates {
<#
.SYNOPSIS
	Removes SCCM cached updates.
.DESCRIPTION
	Removes specified SCCM cached update if it's not needed anymore.
.PARAMETER
.EXAMPLE
.NOTES
.LINK
#>

	## Get list of updates
	$CM_Updates = Get-WmiObject -Namespace root\ccm\SoftwareUpdates\UpdatesStore -Query 'SELECT UniqueID,Title,Status FROM CCM_UpdateStatus'

	## Check if cached updates are not needed and delete them
	ForEach ($Update in $CM_Updates) {

		# Show Progrss bar
		If ($CM_Updates.Count -ne $null) {
			$ProgressCounter++
			Write-Progress -Activity 'Processing Updates' -CurrentOperation $Update.Title -PercentComplete (($ProgressCounter / $CM_Updates.Count) * 100)
		}

		# Check if update is already installed
		If ($Update.Status -eq "Installed") {

			# Call Remove-CacheItem function
			Remove-CacheItem -CacheTD $Update.UniqueID -CacheN $Update.Title
		}
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

## Call Remove-CachedApplications function
Remove-CachedApplications

## Call Remove-CachedApplications function
Remove-CachedPackages

## Call Remove-CachedApplications function
Remove-CachedUpdates

## Get Result sort it and build Result Object
$Result =  $Global:Result | Sort-Object Size`(MB`) -Descending

# Calculate total deleted size
$TotalDeletedSize = $Result | Measure-Object -Property Size`(MB`) -Sum | Select -ExpandProperty Sum

# If $TotalDeletedSize is zero write that nothing could be deleted
If ($TotalDeletedSize -eq $null) {
	$TotalDeletedSize = "Nothing to Delete!"
}

# Build Result Object
$ResultProps = [ordered]@{
	'DeletionDate' = $Date
	'Name' = 'Total Size of Items Deleted in MB:'
	'ID' = ''
	'Location' = ''
	'Size(MB)' = ''
	'TotalDeleted(MB)' = '{0:N2}' -f $TotalDeletedSize
}

# Add total items deleted to result object
$Result += New-Object PSObject -Property $ResultProps

## Write Result Object to csv file (append)
$Result | Export-Csv -Path $ResultCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append -Force

## Write Result to console
$Result | Format-Table Name,TotalDeleted`(MB`)

## Let the user know we are finished
Write-Host ""
Write-Host "Processing Finished!" -BackgroundColor Green -ForegroundColor White

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
