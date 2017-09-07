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
        This SQL Function is used to get the Computer Service Pack Short Name by ResourceID.
    .DESCRIPTION
        This SQL Function is used to get the Computer Service Pack Short Name by ResourceID.
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

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
