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
    This SQL Query is used to get the Compliance and Actual Values of a Configuration Baseline Result.
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

SELECT Distinct
    SYS.Name0 AS 'Computer Name',
    SNames.StateName AS 'Compliance State',
    CCS.LastComplianceMessageTime as 'Last Compliance Evaluation',
    SYS.User_Name0 as 'User Name',
    OS.Caption0 as 'Operating System',
    OS.InstallDate0 as 'Install Date',
    STATUS.LastHWScan as 'Last HW Scan',
    COMP.Model0 as 'Model',
    CI.CIVersion AS 'Baseline Content Version',
    CSCSD.CurrentValue,

    /* IMPORTANT! YOU NEED TO ENABLE THE COMPANY FIELD GATHERING FOR SYSTEM DISCOVERY OTHERWISE THIS COLUMN IS NOT AVAILABLE */
    SYS.Company0 as 'Company'
    --( SELECT [CM_Tools].[dbo].[ufn_GetCompany_by_ResourceID]([sys].[ResourceID]) ) AS [Company]
FROM
    v_BaselineTargetedComputers BTC
    INNER JOIN fn_rbac_R_System(@UserSIDs) as SYS ON SYS.ResourceID = BTC.ResourceID
    LEFT OUTER JOIN v_ClientCollectionMembers c ON C.ResourceID = SYS.ResourceID
    INNER JOIN v_GS_COMPUTER_SYSTEM COMP on COMP.ResourceID = BTC.ResourceID
    INNER JOIN v_GS_OPERATING_SYSTEM OS on OS.ResourceID = SYS.ResourceID
    INNER JOIN v_ConfigurationItems CI ON CI.CI_ID = BTC.CI_ID
    INNER JOIN v_CICurrentComplianceStatus CCS ON CCS.CI_ID = CI.CI_ID AND CCS.ResourceID = BTC.ResourceID
    INNER JOIN v_CIComplianceStatusComplianceDetail CSCD ON CSCD.CI_ID = CI.CI_ID
    INNER JOIN v_CICurrentSettingsComplianceStatusDetail CSCSD ON CSCSD.CI_ID = CSCD.Setting_CI_ID
    INNER JOIN v_LocalizedCIProperties_SiteLoc CIProp ON CIProp.CI_ID = CI.CI_ID
    INNER JOIN v_StateNames SNames ON CCS.ComplianceState = SNames.StateID
    LEFT OUTER JOIN v_GS_WORKSTATION_STATUS STATUS on STATUS.ResourceID=SYS.ResourceID
    LEFT OUTER JOIN v_R_User USR on USR.User_Name0 = SYS.User_Name0
WHERE
    CIProp.DisplayName = @Baseline
        AND SNames.TopicType = 401
        AND C.CollectionID= @CollID
ORDER BY SNames.StateName

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
