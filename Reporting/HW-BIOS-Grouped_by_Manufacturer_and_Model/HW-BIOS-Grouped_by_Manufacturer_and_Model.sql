
/*
*********************************************************************************************************
* Requires        | SQL, company AD attribute, Wi32_Computer_System_Product WMI class gathering         *
* ===================================================================================================== *
* Created by      |    Date    | Comments                                                               *
* _____________________________________________________________________________________________________ *
* Octavian Cordos | 2018-01-18 | First version                                                          *
* Ioan Popovici   |            |                                                                        *
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

/* For testing only */
--DECLARE @UserSIDs VARCHAR(16);
--SELECT @UserSIDs = 'disabled';
--DECLARE @CollectionID VARCHAR(16);
--SELECT @CollectionID = 'WT10000A';
--DECLARE @ExcludeVirtualMachines VARCHAR(5);
--SELECT @ExcludeVirtualMachines = 'NO';

SELECT
    
    /* IMPORTANT! YOU NEED TO ENABLE THE COMPANY FIELD GATHERING FOR SYSTEM DISCOVERY OTHERWISE THIS COLUMN IS NOT AVAILABLE */
    Systems.Company0 as 'Company',
    --( SELECT [CM_Tools].[dbo].[ufn_GetCompany_by_ResourceID]([Systems].[ResourceID]) ) AS [Company],
    Systems.Manufacturer0 AS [Manufacturer],
    CASE 
        WHEN Systems.Model0 LIKE '10AA%' THEN 'ThinkCentre M93p'
        WHEN Systems.Model0 LIKE '10AB%' THEN 'ThinkCentre M93p'
        WHEN Systems.Model0 LIKE '10AE%' THEN 'ThinkCentre M93z'
        WHEN Systems.Model0 LIKE '10FLS1TJ%' THEN 'ThinkCentre M900'
        WHEN Product.Version0 = 'Lenovo Product' THEN ('Unknown ' + Systems.Model0)
        WHEN Systems.Manufacturer0 = 'LENOVO' THEN Product.Version0
	    ELSE Systems.Model0
    END AS Model,
    Systems.Name0 AS [Computer Name],
    Systems.UserName0 AS [User Name],   
    BIOS.Name0 AS [BIOS Name],
    BIOS.Version0 AS [BIOS Version],
    BIOS.SMBIOSBIOSVersion0 AS [SMBIOS Version],
    BIOS.SerialNumber0 AS [BIOS Serial Number],
    OperatingSystem.Caption0 AS [Operating System],
    OperatingSystem.CSDVersion0 AS [OS Service Pack],
    OperatingSystem.Version0 AS [OS Version],
    OperatingSystem.InstallDate0 AS [OS Installed Date]
FROM dbo.v_GS_COMPUTER_SYSTEM AS Systems
    JOIN dbo.v_GS_OPERATING_SYSTEM OperatingSystem ON OperatingSystem.ResourceID = Systems.ResourceID
    JOIN dbo.v_ClientCollectionMembers AS Collections ON Collections.ResourceID = Systems.ResourceID
    JOIN dbo.v_GS_PC_BIOS BIOS ON BIOS.ResourceID = Systems.ResourceID
    JOIN dbo.v_GS_COMPUTER_SYSTEM_PRODUCT AS Product ON Product.ResourceID = Systems.ResourceID
WHERE Collections.CollectionID = @CollectionID
    AND
        Systems.Model0 NOT LIKE (
		    CASE @ExcludeVirtualMachines
			    WHEN 'YES' THEN '%Virtual%'
			    ELSE ''
		    END
	    )
ORDER BY 
    Model, 
    [BIOS Name], 
    [BIOS Version]

/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/