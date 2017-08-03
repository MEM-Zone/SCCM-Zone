USE [CM_Tools]
GO
/****** Object:  UserDefinedFunction [dbo].[ufn_GetCompany_by_ResourceID]    Script Date: 2017-08-03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[ufn_GetCompany_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = (
            SELECT TOP 1
                CASE
                   WHEN ([ou].[System_OU_Name0] LIKE '%ICC%' OR [rn].[Resource_Names0] LIKE '%ICC%') THEN 'ICC'
                   WHEN ([ou].[System_OU_Name0] LIKE '%WTO%' OR [rn].[Resource_Names0] LIKE '%WTO%') THEN 'WTO'
                   WHEN ([ou].[System_OU_Name0] LIKE '%WMO%' OR [rn].[Resource_Names0] LIKE '%WMO%') THEN 'WMO'
                   WHEN ([ou].[System_OU_Name0] LIKE '%WIPO%' OR [rn].[Resource_Names0] LIKE '%WIPO%') THEN 'WIPO'
                   WHEN ([ou].[System_OU_Name0] LIKE '%UNJSPF%' OR [rn].[Resource_Names0] LIKE '%UNJSPF%') THEN 'UNJSPF'
                   WHEN ([ou].[System_OU_Name0] LIKE '%OHCHR%' OR [rn].[Resource_Names0] LIKE '%OHCHR%') THEN 'OHCHR'
                   WHEN ([ou].[System_OU_Name0] LIKE '%SVC%' OR [rn].[Resource_Names0] LIKE '%SVC%') THEN 'SVC'
                   ELSE '0-UNKNOWN'
                END
            FROM [CM_A01].[dbo].[v_RA_System_SystemOUName] [ou]
            LEFT JOIN [CM_A01].[dbo].[v_RA_System_ResourceNames] [rn] ON @pResourceID = [rn].[ResourceID]
            WHERE [ou].[System_OU_Name0] IS NOT NULL
                AND [ou].[ResourceID] = @pResourceID
        )
        RETURN @RET;
    END;
