/*
*********************************************************************************************************
* Created by Ioan Popovici   | Requirements: CM_Tools Database                                          *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2015-08-18 | v1.0     | First version                                                 *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This SQL Function is used to get the Computer Default Gateway by ResourceID.
.DESCRIPTION
    This SQL Function is used to get the Computer Default Gateway by ResourceID.
.EXAMPLE
    Run the code in SQL Server Management Studio.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
*/

/*##=============================================*/
/*## QUERY BODY
/*##=============================================*/
/* #region QueryBody */

USE [CM_Tools]
GO

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

    /* #endregion */
    /*##=============================================*/
    /*## END QUERY BODY                              */
    /*##=============================================*/
