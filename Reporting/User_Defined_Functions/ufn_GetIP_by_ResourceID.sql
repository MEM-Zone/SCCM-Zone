/*
*********************************************************************************************************
* Created by Ioan Popovici, 2015-08-18       | Requirements: CM_Tools Database                          *
* ======================================================================================================*
* Modified by                   |    Date    | Revision | Comments                                      *
*_______________________________________________________________________________________________________*
* Ioan Popovici                 | 2015-09-05 | v1.0     | First version                                 *
*-------------------------------------------------------------------------------------------------------*
*                                                                                                       *
*********************************************************************************************************

    .SYNOPSIS
        This SQL Function is used to get the Computer IP by ResourceID.
    .DESCRIPTION
        This SQL Function is used to get the Computer IP by ResourceID.
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

CREATE FUNCTION [dbo].[ufn_GetIP_by_ResourceID](@pResourceID INT)
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

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
