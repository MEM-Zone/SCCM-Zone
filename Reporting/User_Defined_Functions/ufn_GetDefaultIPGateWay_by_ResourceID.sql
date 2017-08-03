USE [CM_Tools]
GO
/****** Object:  UserDefinedFunction [dbo].[ufn_GetDefaultIPGateWay_by_ResourceID]    Script Date: 07/20/2017 17:20:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[ufn_GetDefaultIPGateWay_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = STUFF(
               (
               SELECT N', '+CASE
                                WHEN CHARINDEX(', f', [ne].[DefaultIPGateway0]) > 0
                                THEN SUBSTRING([ne].[DefaultIPGateway0], 0, CHARINDEX(', f', [ne].[DefaultIPGateway0]))
                                WHEN [ne].[DefaultIPGateway0] IS NULL
                                THEN 'Null'
                                ELSE [ne].[DefaultIPGateway0]
                            END
                      FROM [CM_A01].[dbo].[v_GS_NETWORK_ADAPTER_CONFIGURATION] [ne]
                      WHERE [ne].[DefaultIPGateway0] IS NOT NULL
                            AND [ne].[DefaultIPGateway0] NOT LIKE '%::%'
                            AND [ne].[ResourceID] = @pResourceID
                      ORDER BY [ne].[DefaultIPGateway0]
                      FOR XML PATH(N'')
               ), 1, 1, N'');

        RETURN @RET;
    END;
