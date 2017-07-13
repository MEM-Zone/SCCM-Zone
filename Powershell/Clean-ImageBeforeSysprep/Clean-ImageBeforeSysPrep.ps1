<#
*********************************************************************************************************
*                                                                                                       *
*** This PowerShell script is used shrink the image before SysPrep.                                   ***
*                                                                                                       *
*********************************************************************************************************
* Created by Ioan Popovici, 2017-7-10  | Requirements: PowerShell 3.0                                   *
* ======================================================================================================*
* Modified by                   | Date       | Version  | Comments                                      *
*_______________________________________________________________________________________________________*
* Ioan Popovici                 | 2017-7-10 | v1.0     | First version                                  *
* Ioan Popovici                 | 2017-7-10 | v2.0     | Vastly improved                                *
*-------------------------------------------------------------------------------------------------------*
* Execute with: C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe -NoExit -NoProfile -File      *
* Clean-ImageBeforeSysPrep.ps1                                                                          *
* To do:                                                                                                *
* Add error handling.                                                                                   *
* Add better logging.                                                                                   *
*********************************************************************************************************

    .SYNOPSIS
        This PowerShell script is used shrink the image before SysPrep.
    .DESCRIPTION
        This PowerShell script is used shrink the image before SysPrep by removing volume caches, update backups and update caches.
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration

[String]$mName = "."
[String]$regExPattern =  '(Windows\ (?:7|8\.1|8|Server\ (?:2008\ R2|2008|2012\ R2|2012|2016)))'
[String]$mOS = (Get-WmiObject -class Win32_OperatingSystem -ComputerName $mName | Select-Object Caption | `
    Select-String -AllMatches -Pattern $regExPattern | Select-Object -ExpandProperty Matches).Value
[String]$regRootPath = 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
[Array]$regPaths = Get-ChildItem -Path $regRootPath | Select-Object -ExpandProperty Name
[String]$regSageSet = '5432'
[String]$regName = 'StateFlags'+$regSageSet
[String]$regValue = '00000002'
[String]$regType = 'DWORD'

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Start-Cleanup
Function Start-Cleanup {
    <#
    .SYNOPSIS
        Cleans volume caches, update backups and update caches.
    .DESCRIPTION
        Cleans volume caches, update backups and update caches depending on the selected options.
    .PARAMETER CleanupOptions
        The CleanupOptions depending of what type of cleanup to perform.
    .EXAMPLE
        Start-Cleanup -CleanupOptions ('cCacheCleanup','uCacheCleanup','vCacheCleanup')
    .NOTES
        This is an internal script function and should typically not be called directly.
    .LINK
        http://sccm-zone.com
    #>
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('cOptions')]
        [Array]$CleanupOptions
    )
    Switch ($CleanupOptions) {
        'cCacheCleanup' {
            Start-Process -FilePath DISM.exe /Online /Cleanup-Image /RestoreHealth -Wait
            Start-Process -FilePath DISM.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase -Wait
        }
        'vCacheCleanup' {
            If ($regPaths) {
                Write-Host "Adding $regName to Registry Following Paths:" -BackgroundColor 'DarkGreen'
                $regPaths | ForEach-Object {
                    Write-Host "$_"
                    New-ItemProperty -Path Registry::$_ -Name $regName -Value $regValue -PropertyType $regType -Force | Out-Null
                }
                Start-Process -FilePath CleanMgr.exe /sagerun:$regSageSet -Wait
            }
            Else {
                Write-Host 'Path Not Found! Skipping...' -BackgroundColor 'Red'
            }
        }
        'uCacheCleanup' {
            Stop-Service -Name wuauserv -Verbose | Out-Null
            Remove-Item -Path C:\Windows\SoftwareDistribution\ -Recurse -Force | Out-Null
            Start-Service -Name wuauserv -Verbose
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

If ($mOS) {
    Switch ($mOS) {
        'Windows 7' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'
            Start-Cleanup ('vCacheCleanup','uCacheCleanup')
        }
        'Windows 8' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'
            Start-Cleanup ('vCachesCleanup','uCacheCleanup')
        }
        'Windows 8.1' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'
            Start-Cleanup ('cCacheCleanup','vCachesCleanup','uCacheCleanup')
        }
        'Windows 10' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'
            Start-Cleanup ('cCacheCleanup','vCachesCleanup','uCacheCleanup')
        }
        'Windows Server 2008' {
            Write-Host "$_ Detected. Unsupported Operating System, Skipping Cleanup! `n" -BackgroundColor 'Red'
        }
        'Windows Server 2008 R2' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'

            Copy-Item -Path 'C:\Windows\winsxs\amd64_microsoft-windows-cleanmgr_31bf3856ad364e35_6.1.7600.16385_none_c9392808773cd7da\cleanmgr.exe' -Destination 'C:\Windows\System32\' | Out-Null
            Copy-Item -Path 'C:\Windows\winsxs\amd64_microsoft-windows-cleanmgr.resources_31bf3856ad364e35_6.1.7600.16385_en-us_b9cb6194b257cc63\cleanmgr.exe.mui' -Destination 'C:\Windows\System32\en-US\' | Out-Null

            Start-Cleanup ('vCachesCleanup','uCacheCleanup')
        }
        'Windows Server 2012' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'
            Start-Cleanup ('cCacheCleanup','uCacheCleanup')
        }
        'Windows Server 2012 R2' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'
            Start-Cleanup ('cCacheCleanup','uCacheCleanup')
        }
        'Windows Server 2016' {
            Write-Host "$_ Detected. Starting Cleanup... `n" -BackgroundColor 'Blue'
            Start-Cleanup ('cCacheCleanup','uCacheCleanup')
        }
        Default {
            Write-Host "$_ Detected. Unknown Operating System, Skipping Cleanup! `n" -BackgroundColor 'Red'
        }
    }
}
Else {
    Write-Host "Unknown Operating System, Skipping Cleanup! `n" -BackgroundColor 'Red'
}

#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================
