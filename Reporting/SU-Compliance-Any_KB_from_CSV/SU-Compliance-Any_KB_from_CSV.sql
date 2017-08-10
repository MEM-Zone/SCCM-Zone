
--DECLARE @UserSIDs VARCHAR(16);
--SELECT @UserSIDs = 'disabled';
--DECLARE @CollID VARCHAR(8);
--SET @CollID = 'SMS00001';
--DECLARE @UpdateList Varchar(MAX);
--SET @UpdateList = 'KB4015553,KB4019215,KB4015549,KB4015552,KB4012598,KB4019264,KB4012215,KB4012213,KB4012212,KB4012217,KB4015551,KB4019216,KB4012216,KB4015550,KB4013429,KB4019472,KB4015217,KB4015438,KB4016635,KB4019473,KB4015219,KB4013198,KB4012606,KB4015221,KB4019474,KB4012214,KB4019265,KB4019263,KB4015546,KB4022727,KB4022714,KB4022715,KB4022168,KB4022719,KB4022720,KB4022726,KB4025335,KB4025336,KB4025341,KB4034664,KB4034681,KB4022727,KB4022714,KB4022715,KB4022725,KB4025338,KB4025344,KB4025339,KB4025342,KB4032188,KB4034668,KB4034660,KB4034658,KB4034674'

--SELECT * FROM CM_Tools.dbo.ufn_csv_String_Parser(@UpdateList,',')
--=Join(Parameters!UpdateList.Value,",")

SELECT
    SYS.Name0,
    CASE
        WHEN HE.HotfixID0 IN (SELECT * FROM CM_Tools.dbo.ufn_csv_String_Parser(@UpdateList,',')) THEN 'TRUE'
        ELSE 'FALSE'
    END AS 'Compliant',
    HE.HotfixID0,
    HE.ResourceID
INTO #TMP_RawCompliance
FROM fn_rbac_GS_SYSTEM (@UserSIDs) [SYS]
LEFT JOIN [dbo].[v_GS_QUICK_FIX_ENGINEERING] AS [HE] ON HE.ResourceID = SYS.ResourceID
WHERE HE.HotFixID0 IN (SELECT * FROM CM_Tools.dbo.ufn_csv_String_Parser(@UpdateList,','))


SELECT DISTINCT
    Name0 AS Machine,
    Compliance = 'Compliant'
INTO #TMP_Compliant
FROM #TMP_RawCompliance
WHERE Compliant = 'TRUE'
ORDER BY
    Name0


SELECT
    SYS.Name0 AS Machine,
    Compliance = 'Non-Compliant'
INTO #TMP_NonCompliant
FROM fn_rbac_GS_SYSTEM (@UserSIDs) AS SYS
LEFT JOIN #TMP_Compliant AS CM ON CM.Machine = SYS.Name0
WHERE CM.Machine IS NULL


SELECT *
INTO #TMP_ProcessedCompliance
FROM #TMP_Compliant
    UNION
SELECT *
FROM #TMP_NonCompliant
ORDER BY
    Compliance ASC

SELECT DISTINCT
    [s].[ResourceID] AS [MachineID],
    (SELECT [CM_Tools].dbo.[ufn_GetCompany_by_ResourceID](s.ResourceID)) AS [Company],
    [r].Resource_Names0 AS [Machine],
    CASE
        WHEN [CM].[Compliance] IS NULL THEN 'Unknown'
        ELSE [CM].[Compliance]
    END AS Compliance,
    CASE
        WHEN ([s].[Client0] = 1) THEN 'Yes'
        ELSE 'No'
    END AS [Client],
    CASE
        WHEN ([s].[Active0] = 1) THEN 'Active'
        WHEN ([s].[Active0] = 0) THEN 'Inactive'
        ELSE 'Unknown'
    END AS [Active],
    CASE
       WHEN (chcs.LastEvaluationHealthy = 1) THEN 'Pass'
       WHEN (chcs.LastEvaluationHealthy = 2) THEN 'Fail'
       ELSE 'Unknown'
    END AS 'Last Evaluation Healthy',
    chcs.LastDDR,
    CASE
        WHEN (DATEDIFF(day, chcs.LastDDR, GETDATE()) <= 14) THEN 'Yes'
        WHEN (DATEDIFF(day, chcs.LastDDR, GETDATE()) >= 14) THEN 'No'
        ELSE 'Unknown'
    END AS 'DDR in the last 14 Days',
    CASE
        WHEN (DATEDIFF(day, os.LastBootUpTime0, GETDATE()) <= 14) THEN 'Yes'
        WHEN (DATEDIFF(day, os.LastBootUpTime0, GETDATE()) >= 14) THEN 'No'
        ELSE 'Unknown'
    END AS 'Rebooted in the last 14 Days',
    CASE
        WHEN ([s].[Client_Version0] IS NULL) THEN 'Unknown'
        ELSE [s].[Client_Version0]
    END AS 'Client Version',
    CASE
        WHEN (MAX([ou].[System_OU_Name0]) IS NULL) THEN 'Unknown'
        ELSE MAX([ou].[System_OU_Name0])
    END AS OUName,
    [os].[Caption0] AS OS
FROM [dbo].[fn_rbac_R_System](@UserSIDs) [s]
    LEFT JOIN #TMP_ProcessedCompliance AS CM ON CM.Machine = s.Name0
    LEFT JOIN [v_RA_System_SystemOUName] AS [ou] ON [s].[ResourceID] = [ou].[ResourceID]
    LEFT JOIN fn_rbac_GS_SYSTEM (@UserSIDs) AS [SYS] ON [s].[ResourceID] = [SYS].[ResourceID]
    LEFT JOIN [v_RA_System_ResourceNames] [r] ON [s].[ResourceID] = [r].[ResourceID]
    LEFT OUTER JOIN dbo.v_GS_OPERATING_SYSTEM AS os ON os.ResourceID = [s].[ResourceID]
    LEFT OUTER JOIN dbo.v_CH_ClientSummary AS chcs ON chcs.ResourceID = [s].[ResourceID]
    JOIN dbo.v_FullCollectionMembership AS fcm ON s.ResourceID = fcm.ResourceID
WHERE fcm.CollectionID = @CollID

GROUP BY
    [CM].[Compliance],
    [r].Resource_Names0,
    [SYS].[SystemRole0],
    [s].[Client0],
    [s].[Active0],
    [s].[Client_Version0],
    [s].[Netbios_Name0],
    [s].[Full_Domain_Name0],
    [s].[ResourceID],
    [chcs].[LastEvaluationHealthy],
    [chcs].[LastDDR],
    [os].[LastBootUpTime0],
    [os].[Caption0]
ORDER BY
    Compliance DESC

DROP TABLE #TMP_RawCompliance
DROP TABLE #TMP_Compliant
DROP TABLE #TMP_NonCompliant
DROP TABLE #TMP_ProcessedCompliance
