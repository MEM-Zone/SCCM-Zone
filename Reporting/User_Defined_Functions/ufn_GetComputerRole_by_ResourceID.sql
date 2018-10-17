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
    This SQL Function is used to get the Computer Role by ResourceID.
.DESCRIPTION
    This SQL Function is used to get the Computer Role by ResourceID by using Machine ROLE Property.
.EXAMPLE
    Run the code in SQL Server Management Studio.
.LINK
    https://SCCM-Zone.com
    https://github.com/Ioan-Popovici/SCCMZone
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

CREATE FUNCTION [dbo].[ufn_GetComputerRole_by_ResourceID](@pResourceID INT)
RETURNS VARCHAR(200)
AS
    BEGIN
        DECLARE @RET VARCHAR(200);
        SELECT @RET = (
            SELECT
                CASE
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

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
