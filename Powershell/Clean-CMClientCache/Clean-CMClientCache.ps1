<#
*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 3.0                                                  *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2015-11-13 | v1.0     | First version                                                 *
* Ioan Popovici | 2015-11-16 | v1.1     | Improved logging                                              *
* Ioan Popovici | 2015-11-17 | v1.2     | Vastly improved                                               *
* Ioan Popovici | 2016-02-03 | v2.0     | Vastly improved                                               *
* Ioan Popovici | 2016-02-04 | v2.1     | Fixed TotalSize decimals                                      *
* Ioan Popovici | 2016-02-19 | v2.2     | EventLog logging support                                      *
* Ioan Popovici | 2016-02-20 | v2.3     | Added check for not downloaded Cache Items, improved logging  *
* Ioan Popovici | 2017-04-26 | v2.4     | Basic error management, formatting cleanup                    *
* Ioan Popovici | 2017-04-26 | v2.5     | Orphaned cache cleanup, null CacheID fix, improved logging    *
* Ioan Popovici | 2017-05-02 | v2.5     | Basic error Management                                        *
* Walker        | 2017-08-08 | v2.6     | Fixed first time run logging bug                              *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This PowerShell Script is used to clean the CCM cache of all unneeded, non persisted content.
.DESCRIPTION
    This PowerShell Script is used to clean the CCM cache of all non persisted content that is not needed anymore.
.EXAMPLE
    Clean-CMClientCache
.NOTES
    It only cleans packages, applications and updates that have a "installed" status, are not persisted, or
    are not needed anymore (Some other checks are performed). Other cache items will NOT be cleaned.
.NOTES
    To Do:
    Not happy, this needs a re-write changing the logic. Now it parses all apps/packages/updates and then looks
    in the cache for it. This search is expensive and optimized, will have to go the other way around if possible.
    Also logging and error handling are crap
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

##*=============================================
##* INITIALIZATION
##*=============================================
#region Initialization

## Cleaning prompt history
CLS

## Global variables
$Global:Result  =@()
$Global:ExclusionList  =@()

## Final Result variable
$Result  =@()

## Initialize progress Counter
$ProgressCounter = 0

## Configure Logging
#  Set log path
$ResultCSV = 'C:\Temp\Clean-CMClientCache.log'

#  Remove previous log it it's more than 500 KB
If (Test-Path $ResultCSV) {
    If ((Get-Item $ResultCSV).Length -gt 500KB) {
        Remove-Item $ResultCSV -Force | Out-Null
    }
}

#  Get log parent path
[String]$ResultPath =  Split-Path $ResultCSV -Parent

#  Create path directory if it does not exist
If ((Test-Path $ResultPath) -eq $False) {
    New-Item -Path $ResultPath -Type Directory | Out-Null
}

## Get the current date
$Date = Get-Date

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
    Writes an event to EventLog.
.DESCRIPTION
    Writes an event to EventLog with a specified source.
.PARAMETER EventLogName
    The EventLog to write to.
.PARAMETER EventLogEntrySource
    The EventLog Entry Source.
.PARAMETER EventLogEntryID
    The EventLog Entry ID.
.PARAMETER EventLogEntryType
    The EventLog Entry Type. (Error | Warning | Information | SuccessAudit | FailureAudit)
.PARAMETER EventLogEntryMessage
    The EventLog Entry Message.
.EXAMPLE
    Write-Log -EventLogName 'Configuration Manager' -EventLogEntrySource 'Script' -EventLogEntryID '1' -EventLogEntryType 'Information' -EventLogEntryMessage 'Clean-CMClientCache was successful'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,Position=0)]
        [Alias('Name')]
        [string]$EventLogName = 'Configuration Manager',
        [Parameter(Mandatory=$false,Position=1)]
        [Alias('Source')]
        [string]$EventLogEntrySource = 'Clean-CMClientCache',
        [Parameter(Mandatory=$false,Position=2)]
        [Alias('ID')]
        [int32]$EventLogEntryID = 1,
        [Parameter(Mandatory=$false,Position=3)]
        [Alias('Type')]
        [string]$EventLogEntryType = 'Information',
        [Parameter(Mandatory=$true,Position=4)]
        [Alias('Message')]
        $EventLogEntryMessage
    )

    ## Initialize log
    If (([System.Diagnostics.EventLog]::Exists($EventLogName) -eq $false) -or ([System.Diagnostics.EventLog]::SourceExists($EventLogEntrySource) -eq $false )) {

        #  Create new log and/or source
        New-EventLog -LogName $EventLogName -Source $EventLogEntrySource

    ## Write to log and console
    }

    #  Convert the Result to string and Write it to the EventLog
    $ResultString = Out-String -InputObject $Result -Width 1000
    Write-EventLog -LogName $EventLogName -Source $EventLogEntrySource -EventId $EventLogEntryID -EntryType $EventLogEntryType -Message $ResultString

    #  Write Result Object to csv file (append)
    $EventLogEntryMessage | Export-Csv -Path $ResultCSV -Delimiter ';' -Encoding UTF8 -NoTypeInformation -Append -Force

    #  Write Result to console
    $EventLogEntryMessage | Format-Table Name,TotalDeleted`(MB`)

}
#endregion


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
    Remove-CacheItem -CacheItemToDelete '{234234234}' -CacheItemName 'Office2003'
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
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

    ## Delete cache item if it's non persisted
    If ($CacheItems.ContentID -contains $CacheItemToDelete) {

        #  Get Cache item location and size
        $CacheItemLocation = $CacheItems | Where {$_.ContentID -Contains $CacheItemToDelete} | Select -ExpandProperty Location
        $CacheItemSize =  Get-ChildItem $CacheItemLocation -Recurse -Force | Measure-Object -Property Length -Sum | Select -ExpandProperty Sum

        #  Check if cache item is downloaded by looking at the size
        If ($CacheItemSize -gt '0.00') {

            #  Connect to resource manager COM object
            $CMObject = New-Object -ComObject 'UIResource.UIResourceMgr'

            #  Using GetCacheInfo method to return cache properties
            $CMCacheObjects = $CMObject.GetCacheInfo()

            #  Delete Cache item
            $CMCacheObjects.GetCacheElements() | Where-Object {$_.ContentID -eq $CacheItemToDelete} |
                ForEach-Object {
                    $CMCacheObjects.DeleteCacheElement($_.CacheElementID)
                    Write-Host 'Deleted: '$CacheItemName -BackgroundColor Red
                }
            #  Build result object
            $ResultProps = [ordered]@{
                'Name' = $CacheItemName
                'ID' = $CacheItemToDelete
                'Location' = $CacheItemLocation
                'Size(MB)' = '{0:N2}' -f ($CacheItemSize / 1MB)
                'Status' = 'Deleted!'
            }

            #  Add items to result object
            $Global:Result  += New-Object PSObject -Property $ResultProps
        }
    }
    Else {
        Write-Host 'Already Deleted:'$CacheItemName '|| ID:'$CacheItemToDelete -BackgroundColor Green
    }
}
#endregion

#region Function Remove-CachedApplications
Function Remove-CachedApplications {
<#
.SYNOPSIS
    Removes cached application.
.DESCRIPTION
    Removes specified SCCM cache application if it's already installed.
.EXAMPLE
    Remove-CachedApplications
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

    ## Get list of applications
    Try {
        $CM_Applications = Get-WmiObject -Namespace root\ccm\ClientSDK -Query 'SELECT * FROM CCM_Application' -ErrorAction Stop
    }
    #  Write to log in case of failure
    Catch {
        Write-Host 'Get SCCM Application List from WMI - Failed!'
    }

    ## Check for installed applications
    Foreach ($Application in $CM_Applications) {

        ## Show progress bar
        If ($CM_Applications.Count -ne $null) {
            $ProgressCounter++
            Write-Progress -Activity 'Processing Applications' -CurrentOperation $Application.FullName -PercentComplete (($ProgressCounter / $CM_Applications.Count) * 100)
        }
        ## Get Application Properties
        $Application.Get()

        ## Enumerate all deployment types for an application
        Foreach ($DeploymentType in $Application.AppDTs) {

            ## Get content ID for specific application deployment type
            $AppType = 'Install',$DeploymentType.Id,$DeploymentType.Revision
            $AppContent = Invoke-WmiMethod -Namespace root\ccm\cimodels -Class CCM_AppDeliveryType -Name GetContentInfo -ArgumentList $AppType

            If ($Application.InstallState -eq 'Installed' -and $Application.IsMachineTarget -and $AppContent.ContentID) {

                ## Call Remove-CacheItem function
                Remove-CacheItem -CacheTD $AppContent.ContentID -CacheN $Application.FullName
            }
            Else {
                ## Add to exclusion list
                $Global:ExclusionList += $AppContent.ContentID
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
.EXAMPLE
    Remove-CachedPackages
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

    ## Get list of packages
    Try {
        $CM_Packages = Get-WmiObject -Namespace root\ccm\ClientSDK -Query 'SELECT PackageID,PackageName,LastRunStatus,RepeatRunBehavior FROM CCM_Program' -ErrorAction Stop
    }
    #  Write to log in case of failure
    Catch {
        Write-Host 'Get SCCM Package List from WMI - Failed!'
    }

    ## Check if any deployed programs in the package need the cached package and add deletion or exemption list for comparison
    ForEach ($Program in $CM_Packages) {

        #  Check if program in the package needs the cached package
        If ($Program.LastRunStatus -eq 'Succeeded' -and $Program.RepeatRunBehavior -ne 'RerunAlways' -and $Program.RepeatRunBehavior -ne 'RerunIfSuccess') {

            #  Add PackageID to Deletion List if not already added
            If ($Program.PackageID -NotIn $PackageIDDeleteTrue) {
                [Array]$PackageIDDeleteTrue += $Program.PackageID
            }

        }
        Else {

            #  Add PackageID to Exemption List if not already added
            If ($Program.PackageID -NotIn $PackageIDDeleteFalse) {
                [Array]$PackageIDDeleteFalse += $Program.PackageID
            }
        }
    }

    ## Parse Deletion List and Remove Package if not in Exemption List
    ForEach ($Package in $PackageIDDeleteTrue) {

        #  Show progress bar
        If ($CM_Packages.Count -ne $null) {
            $ProgressCounter++
            Write-Progress -Activity 'Processing Packages' -CurrentOperation $Package.PackageName -PercentComplete (($ProgressCounter / $CM_Packages.Count) * 100)
            Start-Sleep -Milliseconds 800
        }
        #  Call Remove Function if Package is not in $PackageIDDeleteFalse
        If ($Package -NotIn $PackageIDDeleteFalse) {
            Remove-CacheItem -CacheTD $Package.PackageID -CacheN $Package.PackageName
        }
        Else {
            ## Add to exclusion list
            $Global:ExclusionList += $Package.PackageID
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
.EXAMPLE
    Remove-CachedUpdates
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

    ## Get list of updates
    Try {
        $CM_Updates = Get-WmiObject -Namespace root\ccm\SoftwareUpdates\UpdatesStore -Query 'SELECT UniqueID,Title,Status FROM CCM_UpdateStatus' -ErrorAction Stop
    }
    #  Write to log in case of failure
    Catch {
        Write-Host 'Get SCCM Software Update List from WMI - Failed!'
    }

    ## Check if cached updates are not needed and delete them
    ForEach ($Update in $CM_Updates) {

        #  Show Progress bar
        If ($CM_Updates.Count -ne $null) {
            $ProgressCounter++
            Write-Progress -Activity 'Processing Updates' -CurrentOperation $Update.Title -PercentComplete (($ProgressCounter / $CM_Updates.Count) * 100)
        }

        #  Check if update is already installed
        If ($Update.Status -eq 'Installed') {

            #  Call Remove-CacheItem function
            Remove-CacheItem -CacheTD $Update.UniqueID -CacheN $Update.Title
        }
        Else {
            ## Add to exclusion list
            $Global:ExclusionList += $Update.UniqueID
        }
    }
}
#endregion

#region Function Remove-OrphanedCacheItems
Function Remove-OrphanedCacheItems {
<#
.SYNOPSIS
    Removes SCCM orphaned cached items.
.DESCRIPTION
    Removes SCCM orphaned cache items not found in Applications, Packages or Update WMI Tables.
.EXAMPLE
    Remove-OrphanedCacheItems
.NOTES
    This is an internal script function and should typically not be called directly.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>

    ## Check if cached updates are not needed and delete them
    ForEach ($CacheItem in $CacheItems) {

        #  Show Progress bar
        If ($CacheItems.Count -ne $null) {
            $ProgressCounter++
            Write-Progress -Activity 'Processing Orphaned Cache Items' -CurrentOperation $CacheItem.ContentID -PercentComplete (($ProgressCounter / $CacheItems.Count) * 100)
        }

        #  Check if update is already installed
        If ($Global:ExclusionList -notcontains $CacheItem.ContentID) {

            #  Call Remove-CacheItem function
            Remove-CacheItem -CacheTD $CacheItem.ContentID -CacheN 'Orphaned Cache Item'
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

## Get list of all non persisted content in CCMCache, only this content will be removed
Try {
    $CacheItems = Get-WmiObject -Namespace root\ccm\SoftMgmtAgent -Query 'SELECT ContentID,Location FROM CacheInfoEx WHERE PersistInCache != 1' -ErrorAction Stop
}
#  Write to log in case of failure
Catch {
    Write-Host 'Getting SCCM Cache Info from WMI - Failed! Check if SCCM Client is Installed!'
}

## Call Remove-CachedApplications function
Remove-CachedApplications

## Call Remove-CachedApplications function
Remove-CachedPackages

## Call Remove-CachedApplications function
Remove-CachedUpdates

## Call Remove-OrphanedCacheItems function
Remove-OrphanedCacheItems

## Get Result sort it and build Result Object
$Result =  $Global:Result | Sort-Object Size`(MB`) -Descending

#  Calculate total deleted size
$TotalDeletedSize = $Result | Measure-Object -Property Size`(MB`) -Sum | Select -ExpandProperty Sum

#  If $TotalDeletedSize is zero write that nothing could be deleted
If ($TotalDeletedSize -eq $null -or $TotalDeletedSize -eq '0.00') {
    $TotalDeletedSize = 'Nothing to Delete!'
}
Else {
    $TotalDeletedSize = '{0:N2}' -f $TotalDeletedSize
    }

#  Build Result Object
$ResultProps = [ordered]@{
    'Name' = 'Total Size of Items Deleted in MB: '+$TotalDeletedSize
    'ID' = 'N/A'
    'Location' = 'N/A'
    'Size(MB)' = 'N/A'
    'Status' = ' ***** Last Run Date: '+$Date+' *****'
}

#  Add total items deleted to result object
$Result += New-Object PSObject -Property $ResultProps

## Write to log and console
Write-Log -Message $Result

## Let the user know we are finished
Write-Host 'Processing Finished!' -BackgroundColor Green -ForegroundColor White

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
