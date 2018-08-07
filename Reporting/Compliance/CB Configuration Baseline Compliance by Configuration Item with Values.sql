/*
*********************************************************************************************************
* Created by Ioan Popovici   | Requires SQL, company AD attribute gathering, Configuration Baseline     *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2017-09-22 | v1.0     | First version                                                 *
* Ioan Popovici | 2018-01-17 | v1.1     | Fixed: Actual Value is NULL compliance is not displayed       *
* Ioan Popovici | 2018-06-21 | v2.0     | Completly re-written to optimize speed                        *
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

/* Remove previous temporary table if exists */
IF OBJECT_ID (N'TempDB.DBO.#CIComplianceStatusDetails') IS NOT NULL
    BEGIN
        DROP TABLE #CIComplianceStatusDetails;
    END;

/* Get Configuration Item Current Value */
WITH CTE
    AS (
        SELECT
            CIComplianceStatusDetails.ResourceID,
            CIComplianceStatusDetails.CurrentValue,
            CIComplianceStatusDetails.LastComplianceMessageTime,
            RN = ROW_NUMBER()
            OVER (
                PARTITION BY CIComplianceStatusDetails.Netbios_Name0 ORDER BY CIComplianceStatusDetails.LastComplianceMessageTime DESC
            )
        FROM dbo.fn_rbac_CICurrentSettingsComplianceStatusDetail(@UserSIDs) AS CIComplianceStatusDetails
        WHERE CIComplianceStatusDetails.CI_ID
            IN (
                SELECT ReferencedCI_ID
                FROM dbo.fn_rbac_CIRelation_All(@UserSIDs)
                WHERE CI_ID = @BaselineID
                    AND RelationType NOT IN ('7', '0') --Exlude itself and no relation
            )
    )
SELECT
    ResourceID,
    CurrentValue,
    LastComplianceMessageTime
INTO #CIComplianceStatusDetails
FROM CTE
WHERE RN = 1
ORDER BY ResourceID;

SELECT
    CIComplianceState.DisplayName,
    CIComplianceState.ComplianceStateName AS ComplianceState,
    Computer.Name0 AS DeviceName,
    Users.User_Name0 AS UserName,
    OperatingSystem.Caption0 AS OperatingSystem,
    Computer.Model0 AS [Model],
    CIComplianceState.CIVersion,
    CIComplianceStatusDetails.CurrentValue,
    CIComplianceStatusDetails.LastComplianceMessageTime AS LastEvaluation,
    ComputerStatus.LastHWScan AS LastHWScan
FROM v_BaselineTargetedComputers Baseline
    JOIN dbo.fn_rbac_R_System(@UserSIDs) AS Users ON Users.ResourceID = Baseline.ResourceID
    JOIN v_ClientCollectionMembers AS Collections ON Collections.ResourceID = Users.ResourceID
    JOIN v_GS_COMPUTER_SYSTEM Computer ON Computer.ResourceID = Users.ResourceID
    JOIN v_GS_OPERATING_SYSTEM AS OperatingSystem ON OperatingSystem.ResourceID = Users.ResourceID
    JOIN dbo.fn_rbac_CIRelation_All(@UserSIDs) AS BaselineCIs ON BaselineCIs.CI_ID = @BaselineID
        AND BaselineCIs.RelationType NOT IN ('7', '0') --Exlude itself and no relation
    JOIN dbo.fn_rbac_ConfigurationItems(@UserSIDs) CIInfo ON CIInfo.CI_ID = BaselineCIs.ReferencedCI_ID
    JOIN dbo.fn_rbac_ListCI_ComplianceState(@LocaleID, @UserSIDs) AS CIComplianceState ON CIComplianceState.CI_ID = BaselineCIs.ReferencedCI_ID
        AND CIComplianceState.ResourceID = Users.ResourceID
    JOIN v_GS_WORKSTATION_STATUS AS ComputerStatus ON ComputerStatus.ResourceID = Users.ResourceID
    LEFT OUTER JOIN #CIComplianceStatusDetails AS CIComplianceStatusDetails ON CIComplianceStatusDetails.ResourceID = Baseline.ResourceID
WHERE Collections.CollectionID = @CollectionID
    AND Baseline.CI_ID = @BaselineID
ORDER BY
	DisplayName,
	ComplianceStateName

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/