/*
*********************************************************************************************************
* Created by Ioan Popovici   | Requires SQL, company AD attribute gathering, Configuration Baseline     *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2017-09-22 | v1.0     | First version                                                 *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This SQL Query is used to get the Compliance of a Configuration Baseline.
.DESCRIPTION
    This SQL Query is used to get the Compliance of a Configuration Baseline Result.
.NOTES
    Part of a report should not be run separately.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/
/* #region QueryBody */

SELECT DISTINCT
    Computer.Name0 AS 'Computer Name',
    StateName.StateName AS 'Compliance State',
    ComplianceStatus.LastComplianceMessageTime as 'Last Compliance Evaluation',
    Users.User_Name0 AS 'User Name',
	OperatingSystem.Caption0 AS 'Operating System',
	OperatingSystem.InstallDate0 AS 'Install Date',
    ComputerStatus.LastHWScan AS 'Last HW Scan',
	Computer.Model0 AS 'Model',
    ConfigurationItem.CIVersion AS 'Baseline Content Version'

    /* IMPORTANT! YOU NEED TO ENABLE THE COMPANY FIELD GATHERING FOR SYSTEM DISCOVERY OTHERWISE THIS COLUMN IS NOT AVAILABLE */
    -- ,Users.Company0 AS 'Company'

    /* CUSTOM FUNCTION LEAVE DISABLED */
    --( SELECT [CM_Tools].[dbo].[ufn_GetCompany_by_ResourceID]([Users].[ResourceID]) ) AS [Company]
FROM
    v_BaselineTargetedComputers Baseline
	INNER JOIN fn_rbac_R_System(@UserSIDs) AS Users ON Users.ResourceID = Baseline.ResourceID
    LEFT OUTER JOIN v_ClientCollectionMembers AS Collections ON Collections.ResourceID = Users.ResourceID
    INNER JOIN v_GS_COMPUTER_SYSTEM Computer ON Computer.ResourceID = Baseline.ResourceID
	INNER JOIN v_GS_OPERATING_SYSTEM AS OperatingSystem ON OperatingSystem.ResourceID = Users.ResourceID
    INNER JOIN v_ConfigurationItems AS ConfigurationItem ON ConfigurationItem.CI_ID = Baseline.CI_ID
    INNER JOIN v_CICurrentComplianceStatus AS ComplianceStatus ON ComplianceStatus.CI_ID = ConfigurationItem.CI_ID AND ComplianceStatus.ResourceID = Baseline.ResourceID
    INNER JOIN v_LocalizedCIProperties_SiteLoc AS BaselineProperties ON BaselineProperties.CI_ID = ConfigurationItem.CI_ID
    INNER JOIN v_StateNames AS StateName ON ComplianceStatus.ComplianceState = StateName.StateID
    LEFT OUTER JOIN v_GS_WORKSTATION_STATUS AS ComputerStatus ON ComputerStatus.ResourceID = Users.ResourceID
WHERE
    BaselineProperties.DisplayName = @BaselineName
        AND StateName.TopicType = 401
        AND Collections.CollectionID= @CollectionID
ORDER BY StateName.StateName

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/