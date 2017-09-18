<#
*********************************************************************************************************
* Created by Ioan Popovici   | Requires PowerShell 3.0                                                  *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2015-09-06 | v1.0     | First version                                                 *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************
#>

Function Set-PageFile
{
<#
    .SYNOPSIS
        Set-PageFile is an advanced function which can be used to adjust virtual memory page file size.
    .DESCRIPTION
        Set-PageFile is an advanced function which can be used to adjust virtual memory page file size.
    .PARAMETER  <InitialSize>
        Setting the paging file's initial size.
    .PARAMETER  <MaximumSize>
        Setting the paging file's maximum size.
    .PARAMETER  <DriveLetter>
        Specifies the drive letter you want to configure.
    .PARAMETER  <SystemManagedSize>
        Allow Windows to manage page files on this computer.
    .PARAMETER  <None>
        Disable page files setting.
    .PARAMETER  <Reboot>
        Reboot the computer so that configuration changes take effect.
    .PARAMETER  <AutoConfigure>
        Automatically configure the initial size and maximumsize.
    .EXAMPLE
        C:\PS> .\Set-PageFile.ps1 -InitialSize 1024 -MaximumSize 2048 -DriveLetter 'C', 'D'

        Setting the page file size on "C:" - Successful.
        Setting the page file size on "D:" - Successful.

        Name            InitialSize(MB) MaximumSize(MB)
        ----            --------------- ---------------
        C:\pagefile.sys            1024            2048
        D:\pagefile.sys            1024            2048
        E:\pagefile.sys            2048            2048
    .LINK
        Get-WmiObject
        http://technet.microsoft.com/library/hh849824.aspx
    .LINK
        https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
#>
    Param
    (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "SetPageFileSize")]
        [Alias('is')]
        [Int32]$InitialSize = "2047",
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "SetPageFileSize")]
        [Alias('ms')]
        [Int32]$MaximumSize = 2048,
        [Parameter(Mandatory = $false, Position = 2)]
        [Alias('dl')]
        [String[]]$DriveLetter = "S:",
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = "None")]
        [Switch]$None,
        [Parameter(Mandatory = $true, Position = 4, ParameterSetName = "SystemManagedSize")]
        [Switch]$SystemManagedSize,
        [Parameter(Mandatory = $false, Position = 5)]
        [Switch]$Reboot,
        [Parameter(Mandatory = $true, Position = 6, ParameterSetName = "AutoConfigure")]
        [Alias('auto')]
        [Switch]$AutoConfigure,
        [Parameter(Mandatory = $true, Position = 7, ParameterSetName = "AutoVisma")]
        [Alias('autov')]
        [Switch]$AutoVisma
    )
    begin
    {
        Filter Set-PageFileSize
        {
            #Param($DL, $InitialSize, $MaximumSize)

            <#
                The AutomaticManagedPagefile property determines whether the system managed pagefile is enabled.
                This capability is not available on windows server 2003, XP and lower versions.
                 Only if it is NOT managed by the system and will also allow you to change these.
            #>
            $IsAutomaticManagedPagefile = Get-WmiObject -Class Win32_ComputerSystem | Foreach-Object {$_.AutomaticManagedPagefile}
            if($IsAutomaticManagedPagefile)
            {
                #We must enable all the privileges of the current user before the command makes the WMI call.
                $SystemInfo = Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges
                $SystemInfo.AutomaticManagedPageFile = $false
                [Void]$SystemInfo.Put()
            }

            Write-Verbose "Setting pagefile on $DL"

            #Configuring the page file size
            $PageFile = Get-WmiObject -Class Win32_PageFileSetting -Filter "SettingID = 'pagefile.sys @ $DL'"

            Try
            {
                if($PageFile -ne $null)
                {
                    $PageFile.Delete()
                }
                Set-WmiInstance -Class Win32_PageFileSetting -Arguments @{name = "$DL\pagefile.sys"; InitialSize = 0; MaximumSize = 0} -EnableAllPrivileges |Out-Null
                $PageFile = Get-WmiObject Win32_PageFileSetting -Filter "SettingID = 'pagefile.sys @ $DL'"

                $PageFile.InitialSize = $InitialSize
                $PageFile.MaximumSize = $MaximumSize
                [Void]$PageFile.Put()

                $Result = "Seting the page file size on ""$DL"" drive, Minimum: ""$InitialSize"" Maximum: ""$MaximumSize""   - Successful."
                Write-Host $Result | Write-Log
                Write-Warning "Pagefile configuration changed on computer '$Env:COMPUTERNAME'. The computer must be restarted for the changes to take effect."
            }
            Catch
            {
                $Result = "Seting the page file size on ""$DL"" drive - Failed."
                Write-Error $Result | Write-Log
            }
        }
    }
    process
    {
        Foreach($DL in $DriveLetter)
        {
        if($None)
            {
                $PageFile = Get-WmiObject -Query "Select * From Win32_PageFileSetting Where Name = '$DL\\pagefile.sys'" -EnableAllPrivileges
                if($PageFile -ne $null)
                {
                    $PageFile.Delete()
                    $Result = """$DL"" drive pagefile set None - Successful."
                    Write-Warning $Result | Write-Log
                }
                else
                {
                    $Result = """$DL"" drive pagefile is already set None - Failed."
                    Write-Warning $Result | Write-Log
                }
            }
            elseif($SystemManagedSize)
            {
                $InitialSize = 0
                $MaximumSize = 0

                Set-PageFileSize -DL $DL -InitialSize $InitialSize -MaximumSize $MaximumSize
            }
            elseif($AutoConfigure)
            {
                $InitialSize = 0
                $MaximumSize = 0

                #Getting total physical memory size
                Get-WmiObject -Class Win32_PhysicalMemory | Where-Object{$_.DeviceLocator -ne "SYSTEM ROM"} | `
                ForEach-Object{$TotalPhysicalMemorySize+= [Double]($_.Capacity)/1GB}

                <#
                By default, the minimum size on a 32-bit (x86) system is 1.5 times the amount of physical RAM if physical RAM is less than 1 GB,
                and equal to the amount of physical RAM plus 300 MB if 1 GB or more is installed. The default maximum size is three times the amount of RAM,
                regardless of how much physical RAM is installed.
                #>
                if($TotalPhysicalMemorySize -lt 1)
                {
                    $InitialSize = 1.5*1024
                    $MaximumSize = 1024*3
                    Set-PageFileSize -DL $DL -InitialSize $InitialSize -MaximumSize $MaximumSize
                }
                else
                {
                    $InitialSize = 1024+300
                    $MaximumSize = 1024*3
                    Set-PageFileSize -DL $DL -InitialSize $InitialSize -MaximumSize $MaximumSize
                }
            }
            elseif($AutoVisma)
            {
                 $InitialSize = 0
                $MaximumSize = 0

                #Getting total physical memory size
                Get-WmiObject -Class Win32_PhysicalMemory | Where-Object{$_.DeviceLocator -ne "SYSTEM ROM"} | `
                ForEach-Object{$TotalPhysicalMemorySize+= [Double]($_.Capacity)/1GB}

                <#
                By default, the minimum size on a 32-bit (x86) system is 1.5 times the amount of physical RAM if physical RAM is less than 1 GB,
                and equal to the amount of physical RAM plus 300 MB if 1 GB or more is installed. The default maximum size is three times the amount of RAM,
                regardless of how much physical RAM is installed.
                #>
                if($TotalPhysicalMemorySize -lt 16)
                {
                    $InitialSize = 1024*2
                    $MaximumSize = 1024*4
                    Set-PageFileSize -DL $DL -InitialSize $InitialSize -MaximumSize $MaximumSize
                }
                else
                {
                    $InitialSize = 1024*4
                    $MaximumSize = 1024*8
                    Set-PageFileSize -DL $DL -InitialSize $InitialSize -MaximumSize $MaximumSize
                }
            }
            else
            {
                Set-PageFileSize -DL $DL -InitialSize $InitialSize -MaximumSize $MaximumSize
            }

            if($Reboot)
            {
                Restart-Computer -ComputerName $Env:COMPUTERNAME -Force
            }
        }

        #Get current page file size information
        $GetPageFileInfo = Get-WmiObject -Class Win32_PageFileSetting -EnableAllPrivileges|Select-Object Name, `
            @{Name = "InitialSize(MB)";Expression = {if($_.InitialSize -eq 0){"System Managed"}else{$_.InitialSize}}}, `
            @{Name = "MaximumSize(MB)";Expression = {if($_.MaximumSize -eq 0){"System Managed"}else{$_.MaximumSize}}}| `
            Format-Table -AutoSize
        Write-Output $GetPageFileInfo
        }
    }
}
<#
## PageFile detection

$GetPageFileAuto = Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges | Select-Object AutomaticManagedPagefile

$GetPageFileInfo = Get-WmiObject -Class Win32_PageFileSetting -EnableAllPrivileges | Select-Object Name, InitialSize, MaximumSize

If ($GetPageFileAuto -ne $True) {
    If ($GetPageFileInfo.InitialSize -ge '0') {
        If ($GetPageFileInfo.InitialSizeMB -eq '6144') {
            Write-Output 'Compliant'
        }
        Else { Write-Output 'Non-Compliant' }
    }
    If ($GetPageFileInfo.InitialSize -eq $Null) {
        Write-Output 'No PageFile Detected'
    }
}
Else { Write-Output 'System Managed'}
#>
