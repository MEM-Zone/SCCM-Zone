/* For testing only
DECLARE @UserSIDs VARCHAR(16)= 'Disabled';
DECLARE @CollectionID VARCHAR(16)= 'A010016C';
--DECLARE @CollectionID VARCHAR(16)= 'A010016A';
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
               CM.SMSID AS UserName,
               PKG.Name AS PackageName,
               ADV.ProgramName,
               CD.CollectionName,
               CASE
                   WHEN ADV.AssignedScheduleEnabled = 0
                   THEN 'Available'
                   ELSE 'Required'
               END AS Purpose,
               CAS.LastStateName,
               'MachineName' AS MachineName --Random Value, needed in order to be able to save the report...
        FROM fn_rbac_Advertisement(@UserSIDs) ADV
             INNER JOIN fn_rbac_Package2(@UserSIDs) PKG ON ADV.PackageID = PKG.PackageID
             LEFT JOIN fn_rbac_ClientAdvertisementStatus(@UserSIDs) CAS ON CAS.AdvertisementID = ADV.AdvertisementID
             INNER JOIN vClassicDeployments CD ON CD.DeploymentID = ADV.AdvertisementID
             INNER JOIN fn_rbac_FullCollectionMembership(@UserSIDs) CM ON ADV.CollectionID = CM.CollectionID
                                                                          AND ResourceType = 4
        WHERE CM.SMSID IN
(
    SELECT SMSID
    FROM #CollectionMembers
    WHERE ResourceType = 4 --Ony Users
)
        ORDER BY UserName,
                 PackageName,
                 ProgramName,
                 CollectionName,
                 Purpose,
                 LastStateName;
    END;

--Device collection query
IF @CollectionType = 2
    BEGIN
        SELECT DISTINCT
               SYS.Netbios_Name0 AS MachineName,
               PKG.Name AS PackageName,
               ADV.ProgramName,
               DS.CollectionName,
               CASE
                   WHEN DS.Purpose = 0
                   THEN 'Available'
                   ELSE 'Required'
               END AS Purpose,
               CAS.LastStateName
        FROM fn_rbac_Advertisement(@UserSIDs) ADV
             JOIN fn_rbac_Package2(@UserSIDs) PKG ON ADV.PackageID = PKG.PackageID
             JOIN fn_rbac_ClientAdvertisementStatus(@UserSIDs) CAS ON CAS.AdvertisementID = ADV.AdvertisementID
             JOIN fn_rbac_R_System(@UserSIDs) SYS ON CAS.ResourceID = SYS.ResourceID
             JOIN vClassicDeployments DS ON ADV.CollectionID = DS.CollectionID
                                            AND ADV.ProgramName != '*' --Only Programs
        WHERE SYS.Netbios_Name0 IN
(
    SELECT Name
    FROM #CollectionMembers
    WHERE ResourceType = 5 --Only Devices
)
        ORDER BY MachineName,
                 PackageName,
                 ProgramName,
                 CollectionName,
                 Purpose,
                 LastStateName;
    END;

--Remove  temporary table
DROP TABLE #CollectionMembers;