--DECLARE @UserSIDs VARCHAR(16)= 'Disabled';
--DECLARE @CollectionID VARCHAR(16)= 'A010016A';
--DECLARE @CollectionID VARCHAR(16)= 'A010016B';
--DECLARE @Locale INT= 2;
DECLARE @SelectBy VARCHAR(16);
DECLARE @CollectionType VARCHAR(16);

SELECT @SelectBy = ResourceID
FROM Fn_Rbac_FullCollectionMembership(@UserSIDs) AS CM
WHERE CM.CollectionID = @CollectionID
      AND CM.ResourceType = 5; --Device collection

IF @SelectBy > 0
    SET @CollectionType = 'DeviceCollection';
    ELSE
SET @CollectionType = 'UserCollection';

--Remove previous temporary table if exists
IF OBJECT_ID(N'TempDB.DBO.#CollectionMembers') IS NOT NULL
    BEGIN
        DROP TABLE #CollectionMembers;
    END;

--Get collection members
SELECT *
INTO #CollectionMembers
FROM Fn_Rbac_FullCollectionMembership(@UserSIDs) AS CM
WHERE CM.CollectionID = @CollectionID;

--Device collection query
IF @CollectionType = 'DeviceCollection'
    BEGIN
        SELECT DISTINCT
               DS.SoftwareName AS SoftwareName,
               DS.CollectionID,
               DS.CollectionName,
               DS.CI_ID,
               AD.MachineName,
               AD.UserName,
               Dbo.Fn_GetAppState(AD.ComplianceState, AD.EnforcementState, CIA.OfferTypeID, 1, AD.DesiredState, AD.IsApplicable) AS EnforcementState
        FROM Fn_Rbac_R_System(@UserSIDs) AS Sd
             INNER JOIN Fn_Rbac_FullCollectionMembership(@UserSIDs) AS CM ON CM.ResourceID = Sd.ResourceID
             INNER JOIN Fn_Rbac_DeploymentSummary(Dbo.FnlShortNameToLCID(@Locale), @UserSIDs) AS DS ON DS.CollectionID = CM.CollectionID
                                                                                                       AND DS.FeatureType = 1
             LEFT JOIN Fn_Rbac_AppIntentAssetData(@UserSIDs) AS Ad ON AD.MachineID = CM.ResourceID
                                                                      AND AD.AssignmentID = DS.AssignmentID
             INNER JOIN Fn_Rbac_CIAssignment(@UserSIDs) AS CIA ON CIA.AssignmentID = DS.AssignmentID
        WHERE Client0 = 1
              AND Sd.Netbios_Name0 IN
(
    SELECT Name
    FROM #CollectionMembers
);
    END;

--User collection query
IF @CollectionType = 'UserCollection'
    BEGIN
        SELECT DISTINCT
               DS.SoftwareName AS SoftwareName,
               DS.CollectionID,
               DS.CollectionName,
               DS.CI_ID,
               AD.MachineName,
               AD.UserName,
               Dbo.Fn_GetAppState(AD.ComplianceState, AD.EnforcementState, CIA.OfferTypeID, 1, AD.DesiredState, AD.IsApplicable) AS EnforcementState
        FROM Fn_Rbac_Collectionexpandedusermembers(@UserSIDs) AS CM
             INNER JOIN Fn_Rbac_R_User(@UserSIDs) AS Ud ON Ud.ResourceID = CM.UserItemKey
             INNER JOIN Fn_Rbac_DeploymentSummary(Dbo.FnlShortNameToLCID(@Locale), @UserSIDs) AS DS ON DS.CollectionID = CM.SiteID
             LEFT JOIN Fn_Rbac_AppIntentAssetData(@UserSIDs) AS Ad ON AD.UserName = Ud.Unique_User_Name0
                                                                      AND AD.AssignmentID = DS.AssignmentID
             INNER JOIN Fn_Rbac_CIAssignment(@UserSIDs) AS CIA ON CIA.AssignmentID = DS.AssignmentID
        WHERE DS.FeatureType = 1
              AND Ud.Unique_User_Name0 IN
(
    SELECT SMSID
    FROM #CollectionMembers
)
        ORDER BY SoftwareName;
    END;

--Remove  temporary table
DROP TABLE #CollectionMembers;