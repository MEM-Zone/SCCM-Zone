USE [CM_Tools]
GO
/****** Object:  UserDefinedFunction [dbo].[ufn_GetOSServicePackShortName_by_ResourceID]    Script Date: 2017-08-03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[ufn_GetOSServicePackShortName_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = (
            SELECT
                CASE
                    WHEN [os].[CSDVersion0] LIKE '%1%' THEN '1'
                    WHEN [os].[CSDVersion0] LIKE '%2%' THEN '2'
                    WHEN [os].[CSDVersion0] LIKE '%3%' THEN '3'
                    WHEN [os].[CSDVersion0] LIKE '%4%' THEN '4'
                    WHEN [os].[CSDVersion0] LIKE '%5%' THEN '5'
                    ELSE '0-UNKNOWN'
                END
            FROM [CM_A01].[dbo].[v_GS_OPERATING_SYSTEM] [os]
            WHERE [os].[CSDVersion0] IS NOT NULL
                AND [os].[ResourceID] = @pResourceID
        )
        RETURN @RET;
    END;
