USE [CM_Tools]
GO
/****** Object:  UserDefinedFunction [dbo].[fnGetIP_by_ResourceID]    Script Date: 07/20/2017 17:21:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[fnGetIP_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = STUFF(
               (
               SELECT N', '+CASE
                                WHEN CHARINDEX(', f', [ne].[IP_Addresses0]) > 0
                                THEN SUBSTRING([ne].[IP_Addresses0], 0, CHARINDEX(', f', [ne].[IP_Addresses0]))
                                WHEN [ne].[IP_Addresses0] IS NULL
                                THEN 'Null'
                                ELSE [ne].[IP_Addresses0]
                            END
                      FROM [CM_A01].[dbo].[v_RA_System_IPAddresses] [ne]
                      WHERE [ne].[IP_Addresses0] IS NOT NULL
                            AND [ne].[IP_Addresses0] NOT LIKE '%::%'
                            AND [ne].[ResourceID] = @pResourceID
                      ORDER BY [ne].[IP_Addresses0]
                      FOR XML PATH(N'')
               ), 1, 1, N'');

        RETURN @RET;
    END;
