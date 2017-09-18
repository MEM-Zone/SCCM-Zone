<#
*********************************************************************************************************
* Created by Ioan Popovici   | 3.0, ADK Windows 10, Windows 8 or higher. Tested on Windows 2012 R2.     *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2017-08-25 | v1.0     | First version                                                 *
* Ioan Popovici | 2017-08-27 | v1.1     | Fixed incorrect DISM version selection                        *
* Ioan Popovici | 2017-08-28 | v1.2     | Improved data input and console output                        *
* Ioan Popovici | 2017-08-29 | v1.3     | Added C:\Windows\System32 path to $Env:Path                   *
* Ioan Popovici | 2017-08-31 | v1.4     | Fixed multiple selection prevention bug                       *
* Ioan Popovici | 2017-09-11 | v1.5     | Fixed $ScriptName variable                                    *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This PowerShell script is used inject packages in a WIM.
.DESCRIPTION
    This PowerShell script is used inject packages in a Windows Image image using PowerShell DISM CommandLets.
.EXAMPLE
    C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe -NoExit -NoProfile -File Update-OfflineWindowsImage.ps1
.NOTES
    Credit for the original VBScript to:
    http://www.catonrug.net/2014/08/slipstream-internet-explorer-11-into-windows-7-sp1-x64-source.html
.NOTES
    To do:
    * Better error handling.
    * Better logging.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM.
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

## Get script path and name
[String]$ScriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
[String]$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
## Set Paths and Image Index
[String]$MountPath = (Join-Path -Path $ScriptPath -ChildPath '\Mount')
[String]$ScratchPath = (Join-Path -Path $ScriptPath -ChildPath '\Scratch')
[String]$UpdatesPath = (Join-Path -Path $ScriptPath -ChildPath '\Updates')
[String]$LogPath = (Join-Path -Path $ScriptPath -ChildPath 'DISM.log')
## Set Environment Path in order to use the latest DISM and System32. This is set for current session only, no need to remove it afterwards.
$Env:Path = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\DISM;C:\Windows\System32'

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================


##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody

## Create Mount directory if it does not exist
If ((Test-Path $MountPath) -eq $False) {
    New-Item -Path $MountPath -Type Directory | Out-Null
}

## Create Scratch directory if it does not exist
If ((Test-Path $ScratchPath) -eq $False) {
    New-Item -Path $ScratchPath -Type Directory | Out-Null
}

## Prompt for WIM to service, don't allow null input
Do {
    [Array]$ImageFileInfo = Get-ChildItem -Path $ScriptPath -Filter '*.wim' | Select-Object -Property `
        @{Label='Name';Expression={($_.Name)}},
        @{Label='Size (GB)';Expression={'{0:N2}' -f ($_.Length / 1GB)}},
        @{Label='Path';Expression={($_.FullName)}} | Out-GridView -PassThru -Title 'Choose image to service. Do not use multiple selection!'
}
While ($ImageFileInfo.Length -eq 0)

#  Set image file name and path, process only the first selection
[String]$ImageFile = ($ImageFileInfo | Select-Object -First 1).Name
[String]$ImagePath = ($ImageFileInfo | Select-Object -First 1).Path

## Prompt for windows version, don't allow null input
Do {
    [Array]$ImageIndexInfo = Get-WindowsImage -ImagePath $ImagePath | Select-Object -First 1 -Property `
        @{Label='Index';Expression={($_.ImageIndex)}},
        @{Label='Name';Expression={($_.ImageName)}},
        @{Label='Description';Expression={($_.ImageDescription)}},
        @{Label='Size (GB)';Expression={'{0:N2}' -f ($_.ImageSize / 1GB)}} | Out-GridView -PassThru -Title 'Choose Windows version. Do not use multiple selection!'
}
While ($ImageIndexInfo.Length -eq 0)

#  Set image name and index, process only the first selection
[String]$ImageName = ($ImageIndexInfo | Select-Object -First 1).Name
[Int]$ImageIndex = ($ImageIndexInfo | Select-Object -First 1).Index

## Prompt for updates to apply , don't allow null input
Do {
    [Array]$Updates = Get-ChildItem -Path $UpdatesPath | Select-Object -Property `
        @{Label='Name';Expression={($_.Name)}},
        @{Label='Size (MB)';Expression={'{0:N2}' -f ($_.Length / 1MB)}},
        @{Label='Path';Expression={($_.FullName)}} | Out-GridView -PassThru -Title 'Choose updates to apply. Use CTRL for multiple selection.'
}
While ($Updates.Length -eq 0)

## Set Backup path and Backup WIM file
[String]$BackupImagePath = ($ImagePath -replace '.{3}$')+'bkp'

Write-Host "`n`n`n`n`n`nBacking up $ImagePath to $BackupImagePath ..." -ForegroundColor 'Yellow' -BackgroundColor 'Black'
$Operation = Copy-Item -Path $ImagePath -Destination $BackupImagePath -Force

## Mount WIM
Write-Host "`nMounting Image..." -ForegroundColor 'Yellow' -BackgroundColor 'Black'
$Operation = Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $MountPath -ScratchDirectory $ScratchPath -LogPath $LogPaths

#  Display mounted image status
$Operation = Get-WindowsImage -Mounted -ScratchDirectory $ScratchPath -LogPath $LogPath | Out-String
Write-Host "`nMounted Image Info:`n $Operation" -ForegroundColor 'Yellow' -BackgroundColor 'Black'

## Add packages to the image
Write-Host "Servicing Image...`n" -ForegroundColor 'Yellow' -BackgroundColor 'Black'
$Updates | ForEach-Object {
    Try {
        Write-Host "Adding $($_.Name) to the Image." -ForegroundColor 'Yellow' -BackgroundColor 'Black'
        $Operation = Add-WindowsPackage -Path $MountPath -PackagePath $_.Path -ScratchDirectory $ScratchPath -LogPath $LogPath
        Write-Host "Adding Package to the Image - Successful!" -ForegroundColor 'Green' -BackgroundColor 'Black'
    }
    Catch {
        Write-Host "Adding Package to the Image - Failed!" -ForegroundColor 'Red' -BackgroundColor 'Black'
    }

}

## Unmount and save servicing changes to the image
Write-Host "`nCommitting Changes and Dismounting Image..." -ForegroundColor 'Yellow' -BackgroundColor 'Black'
$Operation = Dismount-WindowsImage -Path $MountPath -ScratchDirectory $ScratchPath -LogPath $LogPath -Save -CheckIntegrity

##  Set Export path and name
[String]$ExportImagePath =  ($ImagePath -replace '.{4}$')+'_Serviced.wim'

##  Export only selected index also removing unnecessary resource files created during servicing
Write-Host "`nExporting Serviced Image..." -ForegroundColor 'Yellow' -BackgroundColor 'Black'
$Operation = Export-WindowsImage -SourceImagePath $ImagePath -SourceIndex $ImageIndex -DestinationImagePath $ExportImagePath -DestinationName $ImageName -ScratchDirectory $ScratchPath -LogPath $LogPath

## Wait for keypress
Write-Host "Press any key to continue ..."
$WaitforKeyPress = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
