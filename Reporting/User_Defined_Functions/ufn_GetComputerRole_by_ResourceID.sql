USE [CM_Tools]
GO
/****** Object:  UserDefinedFunction [dbo].[ufn_GetComputerRole_by_ResourceID]    Script Date: 2017-08-03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[ufn_GetComputerRole_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = (
            SELECT
                CASE
                    WHEN [cs].Roles0 LIKE '%Domain_Controller%' THEN 'DC'
                    WHEN [cs].Roles0 LIKE '%Domain_Controller%' THEN 'DC'
                    WHEN [cs].Roles0 LIKE '%Workstation%' THEN 'Workstation'
                    WHEN [cs].Roles0 LIKE '%Server%' THEN 'Server'
                    ELSE '0-UNKNOWN'
                END
            FROM [CM_A01].[dbo].[v_GS_COMPUTER_SYSTEM] [cs]
            WHERE [cs].[Caption0] IS NOT NULL
                AND [cs].[ResourceID] = @pResourceID
        )
        RETURN @RET;
    END;
