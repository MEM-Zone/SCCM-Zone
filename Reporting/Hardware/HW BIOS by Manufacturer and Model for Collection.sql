/*
*********************************************************************************************************
* Requires          | SQL, Wi32_Computer_System_Product WMI class gathering                             *
* ===================================================================================================== *
* Modified by       |    Date    | Revision | Comments                                                  *
* _____________________________________________________________________________________________________ *
* Octavian Cordos   | 2018-01-18 | v1.0     | First version                                             *
* Ioan Popovici     | 2018-08-08 | v1.1     | Added localizations, sorting, template, element naming    *
* Ioan Popovici     | 2018-08-21 | v1.2     | Fixed duplicates minor formating fixes                    *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This SQL Query is used to get the Hardware info of a Computer Collection.
.DESCRIPTION
    This SQL Query is used to get the Hardware info of a Computer Collection including BIOS version and Computer model.
.NOTES
    Part of a report should not be run separately.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/

SELECT DISTINCT
    System.Manufacturer0 AS Manufacturer,
    CASE
        WHEN System.Model0 LIKE '10AA%' THEN 'ThinkCentre M93p'
        WHEN System.Model0 LIKE '10AB%' THEN 'ThinkCentre M93p'
        WHEN System.Model0 LIKE '10AE%' THEN 'ThinkCentre M93z'
        WHEN System.Model0 LIKE '10FLS1TJ%' THEN 'ThinkCentre M900'
        WHEN Product.Version0 = 'Lenovo Product' THEN ('Unknown ' + System.Model0)
        WHEN System.Manufacturer0 = 'LENOVO' THEN Product.Version0
        ELSE System.Model0
    END AS Model,
    System.Name0 AS DeviceName,
    System.UserName0 AS UserName,
    BIOS.Name0 AS BIOSName,
    BIOS.Version0 AS BIOSVersion,
    BIOS.SMBIOSBIOSVersion0 AS SMBIOSVersion,
    BIOS.SerialNumber0 AS BIOSSerialNumber,
    OperatingSystem.Caption0 AS OperatingSystem,
    OperatingSystem.CSDVersion0 AS OSServicePack,
    OperatingSystem.Version0 AS OSBuildNumber,
    OperatingSystem.InstallDate0 AS OSInstallDate,
    OSLocalizedNames.Value AS OSVersion
FROM dbo.fn_rbac_GS_COMPUTER_SYSTEM(@UserSIDs) AS System
    JOIN dbo.v_GS_OPERATING_SYSTEM OperatingSystem ON OperatingSystem.ResourceID = System.ResourceID
    LEFT JOIN dbo.vSMS_WindowsServicingStates AS OSServicingStates ON OSServicingStates.Build = OperatingSystem.Version0
    LEFT JOIN vSMS_WindowsServicingLocalizedNames AS OSLocalizedNames ON OSLocalizedNames.Name = OSServicingStates.Name
    JOIN dbo.v_ClientCollectionMembers AS Collections ON Collections.ResourceID = System.ResourceID
    JOIN dbo.v_GS_PC_BIOS BIOS ON BIOS.ResourceID = System.ResourceID
    JOIN dbo.v_GS_COMPUTER_SYSTEM_PRODUCT AS Product ON Product.ResourceID = System.ResourceID
WHERE Collections.CollectionID = @CollectionID
    AND
        System.Model0 NOT LIKE (
            CASE @ExcludeVirtualMachines
                WHEN 'YES' THEN '%Virtual%'
                ELSE ''
            END
        )
ORDER BY
    Model,
    BIOSName,
    BIOSVersion

/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/