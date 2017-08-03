--DECLARE @UserSIDs VARCHAR(16);
--SELECT @UserSIDs = 'disabled';
--DECLARE @CollID VARCHAR(8);
--SET @CollID = 'SMS00001';
--DECLARE @Software VARCHAR(200);
--SET @Software = 'ITSM';

SELECT MAX([ou].[System_OU_Name0]) AS [ouName],
       s.Netbios_Name0,
       [CM_Tools].[dbo].[ufn_GetCompany_by_ResourceID](s.ResourceID) AS Company,
       s.User_Name0,
       s.Operating_System_Name_And0,
       a.DisplayName0,
       a.Version0,
       os.Caption0
FROM dbo.fn_rbac_R_System(@UserSIDs) s
LEFT OUTER JOIN [v_ClientCollectionMembers] c ON [c].[ResourceID] = [s].[ResourceID]
LEFT OUTER JOIN [v_Add_Remove_Programs] a ON [a].[ResourceID] = [s].[ResourceID]
INNER JOIN [dbo].[v_GS_OPERATING_SYSTEM] os ON [s].[ResourceID] = [os].[ResourceID]
LEFT OUTER JOIN [v_RA_System_SystemOUName] ou ON [ou].[ResourceID] = [s].[ResourceID]
WHERE c.CollectionID = @CollID
    AND a.DisplayName0 LIKE '%'+@Software+'%'
GROUP BY
        s.ResourceID,
        s.Netbios_Name0,
        s.User_Name0,
        s.Operating_System_Name_and0,
        a.DisplayName0,
        a.Version0,
        os.Caption0
ORDER BY
        Company,
        a.DisplayName0,
        a.Version0,
        s.Netbios_Name0
