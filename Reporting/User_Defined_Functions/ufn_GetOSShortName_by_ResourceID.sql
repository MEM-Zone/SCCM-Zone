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
        This SQL Function is used to get the Computer OS Short Name by ResourceID.
    .DESCRIPTION
        This SQL Function is used to get the Computer OS Short Name by ResourceID.
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

CREATE FUNCTION [dbo].[ufn_GetOSShortName_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = (
            SELECT
                CASE
                    WHEN os.Caption0 LIKE '%2003%' THEN 'Windows 2003'
                    WHEN os.Caption0 LIKE '%2008R2%' THEN 'Windows 2008 R2'
                    WHEN os.Caption0 LIKE '%2008%' THEN 'Windows 2008'
                    WHEN os.Caption0 LIKE '%2012 R2%' THEN 'Windows 2012 R2'
                    WHEN os.Caption0 LIKE '%2012%' THEN 'Windows 2012'
                    WHEN os.Caption0 LIKE '%2016%' THEN 'Windows 2016'
                    WHEN os.Caption0 LIKE '%XP%' THEN 'Windows XP'
                    WHEN os.Caption0 LIKE '%Windows 7%' THEN 'Windows 7'
                    WHEN os.Caption0 LIKE '%Windows 8%' THEN 'Windows 8'
                    WHEN os.Caption0 LIKE '%Wndows 8.1%' THEN 'Windows 8.1'
                    WHEN os.Caption0 LIKE '%Windows 10%' THEN 'Windows 10'
                    ELSE '0-UNKNOWN'
                END
            FROM [CM_A01].[dbo].[v_GS_OPERATING_SYSTEM] [os]
            WHERE [os].[Caption0] IS NOT NULL
                AND [os].[ResourceID] = @pResourceID
        )
        RETURN @RET;
    END;

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
