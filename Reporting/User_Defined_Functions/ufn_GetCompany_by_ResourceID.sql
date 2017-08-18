USE [CM_Tools]
GO
/****** Object:  UserDefinedFunction [dbo].[ufn_GetCompany_by_ResourceID]    Script Date: 08/10/2017 19:58:10 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

IF EXISTS ( SELECT  [OBJECT_ID]
            FROM    SYS.OBJECTS
            WHERE   NAME = 'ufn_GetCompany_by_ResourceID' )
    DROP FUNCTION dbo.ufn_GetCompany_by_ResourceID;
GO

CREATE FUNCTION [dbo].[ufn_GetCompany_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = (
            SELECT TOP 1
                CASE
                    WHEN ([ou].[System_OU_Name0] LIKE '%WTO%' OR [rn].[Resource_Names0] LIKE '%WTO%') THEN 'WTO'
                    WHEN ([ou].[System_OU_Name0] LIKE '%WMO%' OR [rn].[Resource_Names0] LIKE '%WMO%') THEN 'WMO'
                    WHEN ([ou].[System_OU_Name0] LIKE '%WIPO%' OR [rn].[Resource_Names0] LIKE '%WIPO%') THEN 'WIPO'
                    WHEN ([ou].[System_OU_Name0] LIKE '%UNJSPF%' OR [rn].[Resource_Names0] LIKE '%UNJSPF%') THEN 'UNJSPF'
                    WHEN ([ou].[System_OU_Name0] LIKE '%OHCHR%' OR [rn].[Resource_Names0] LIKE '%OHCHR%') THEN 'OHCHR'
                    WHEN ([ou].[System_OU_Name0] LIKE '%SVC%' OR [rn].[Resource_Names0] LIKE '%SVC%') THEN 'SVC'
                    WHEN ([ou].[System_OU_Name0] LIKE '%ICC%' OR [rn].[Resource_Names0] LIKE '%ICC%') THEN 'ICC'
                    ELSE '0-UNKNOWN'
                END
        FROM [CM_A01].[dbo].[v_RA_System_ResourceNames] [rn]
        LEFT JOIN [CM_A01].[dbo].[v_RA_System_SystemOUName] AS [ou] ON @pResourceID = [ou].[ResourceID]
        WHERE [rn].[Resource_Names0] IS NOT NULL
            AND [rn].[ResourceID] = @pResourceID
        )
        IF @RET IS NULL SET @RET = '0-UNKNOWN'
        RETURN @RET;
    END;
