/* For testing only*
DECLARE @UserSIDs VARCHAR(16)= 'Disabled';
--DECLARE @CollectionID VARCHAR(16)= 'A010016B';
DECLARE @CollectionID VARCHAR(16)= 'A010016A';
DECLARE @Locale INT= 2;
DECLARE @SelectBy VARCHAR(16);
DECLARE @CollectionType VARCHAR(16);
SELECT @SelectBy = ResourceID
FROM fn_rbac_FullCollectionMembership(@UserSIDs) AS CM
WHERE CM.CollectionID = @CollectionID
      AND CM.ResourceType = 5; --Device collection

IF @SelectBy > 0
    SET @CollectionType = 2;
    ELSE
SET @CollectionType = 1;
*/

--Remove previous temporary table if exists
IF OBJECT_ID(N'TempDB.DBO.#CollectionMembers') IS NOT NULL
    BEGIN
        DROP TABLE #CollectionMembers;
    END;

--Get collection members
SELECT *
INTO #CollectionMembers
FROM fn_rbac_FullCollectionMembership(@UserSIDs) AS CM
WHERE CM.CollectionID = @CollectionID
      AND CM.ResourceType IN(4, 5); --Only Users or Devices

--User collection query
IF @CollectionType = 1
    BEGIN
        SELECT DISTINCT
               UD.Unique_User_Name0 AS UserName,
               DS.SoftwareName,
               DS.CollectionName,
               AD.MachineName,
               CASE
                   WHEN CIA.DesiredConfigType = 1
                   THEN 'Install'
                   ELSE 'Remove'
               END AS Purpose,
               AD.UserName AS InstalledBy,
               Dbo.fn_GetAppState(AD.ComplianceState, AD.EnforcementState, CIA.OfferTypeID, 1, AD.DesiredState, AD.IsApplicable) AS EnforcementState
        FROM fn_rbac_CollectionExpandedUserMembers(@UserSIDs) AS CM
             INNER JOIN fn_rbac_R_User(@UserSIDs) AS UD ON UD.ResourceID = CM.UserItemKey
             INNER JOIN fn_rbac_DeploymentSummary(Dbo.FnlShortNameToLCID(@Locale), @UserSIDs) AS DS ON DS.CollectionID = CM.SiteID
             LEFT JOIN fn_rbac_AppIntentAssetData(@UserSIDs) AS Ad ON AD.UserName = UD.Unique_User_Name0
                                                                      AND AD.AssignmentID = DS.AssignmentID
             INNER JOIN fn_rbac_CIAssignment(@UserSIDs) AS CIA ON CIA.AssignmentID = DS.AssignmentID
        WHERE DS.FeatureType = 1
              AND UD.Unique_User_Name0 IN
(
    SELECT SMSID
    FROM #CollectionMembers
    WHERE ResourceType = 4 --Ony Users
)
        ORDER BY UserName,
                 SoftwareName,
                 CollectionName,
                 MachineName,
                 EnforcementState,
                 InstalledBy;
    END;

--Device collection query
IF @CollectionType = 2
    BEGIN
        SELECT DISTINCT
               SD.Netbios_Name0 AS MachineName,
               DS.SoftwareName,
               DS.CollectionName,
               CASE
                   WHEN CIA.DesiredConfigType = 1
                   THEN 'Install'
                   ELSE 'Remove'
               END AS Purpose,
               AD.UserName AS InstalledBy,
               Dbo.fn_GetAppState(AD.ComplianceState, AD.EnforcementState, CIA.OfferTypeID, 1, AD.DesiredState, AD.IsApplicable) AS EnforcementState
        FROM fn_rbac_R_System(@UserSIDs) AS SD
             INNER JOIN fn_rbac_FullCollectionMembership(@UserSIDs) AS CM ON CM.ResourceID = SD.ResourceID
                                                                             AND CM.ResourceType = 5 --Only Devices
             INNER JOIN fn_rbac_DeploymentSummary(Dbo.FnlShortNameToLCID(@Locale), @UserSIDs) AS DS ON DS.CollectionID = CM.CollectionID
                                                                                                       AND DS.FeatureType = 1
             LEFT JOIN fn_rbac_AppIntentAssetData(@UserSIDs) AS AD ON AD.MachineID = CM.ResourceID
                                                                      AND AD.AssignmentID = DS.AssignmentID
             INNER JOIN fn_rbac_CIAssignment(@UserSIDs) AS CIA ON CIA.AssignmentID = DS.AssignmentID
        WHERE Client0 = 1
              AND SD.Netbios_Name0 IN
(
    SELECT Name
    FROM #CollectionMembers
    WHERE ResourceType = 5 --Only Devices
)
        ORDER BY MachineName,
                 SoftwareName,
                 CollectionName,
                 InstalledBy,
                 EnforcementState;
    END;

--Remove  temporary table
DROP TABLE #CollectionMembers;