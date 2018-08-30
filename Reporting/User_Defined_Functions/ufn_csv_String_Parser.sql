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
    This SQL Function is used to parse a CSV string.
.DESCRIPTION
    This SQL Function is used to parse a CSV string and return individual substrings.
.EXAMPLE
    Run the code in SQL Server Management Studio
.NOTES
    Credit to:
    Michelle Ufford http://sqlfool.com.
.LINK
    https://SCCM-Zone.com
    https://github.com/JhonnyTerminus/SCCMZone
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

IF EXISTS
(
    SELECT  [OBJECT_ID]
    FROM    SYS.OBJECTS
    WHERE   NAME = 'ufn_csv_String_Parser'
)
    DROP FUNCTION dbo.ufn_csv_String_Parser;
GO

CREATE FUNCTION [dbo].[ufn_csv_String_Parser]
(
    @pInputString VARCHAR(8000)
,   @pDelimiter CHAR(1)
)
RETURNS @tRET TABLE (StringValue VARCHAR(128))
AS
    BEGIN

        /* Variable declaration */
        DECLARE @pTrimmedInputString VARCHAR(8000);

        /* Trim string input */
        SET @pTrimmedInputString = LTRIM(RTRIM(@pInputString));

        /* Create a recursive CTE to break down the string */
        WITH ParseCTE (StartPos, EndPos)
        AS
        (
            SELECT 1 AS StartPos
                , CHARINDEX(@pDelimiter, @pTrimmedInputString + @pDelimiter) AS EndPos
            UNION ALL
            SELECT EndPos + 1 AS StartPos
                , CHARINDEX(@pDelimiter, @pTrimmedInputString + @pDelimiter , EndPos + 1) AS EndPos
            FROM ParseCTE
            WHERE CHARINDEX(@pDelimiter, @pTrimmedInputString + @pDelimiter, EndPos + 1) <> 0
        )

        /* Insert results into a table */
        INSERT INTO @tRET
        SELECT SUBSTRING(@pTrimmedInputString, StartPos, EndPos - StartPos)
        FROM ParseCTE
        WHERE LEN(LTRIM(RTRIM(SUBSTRING(@pTrimmedInputString, StartPos, EndPos - StartPos)))) > 0
        OPTION (MaxRecursion 8000);

        RETURN;
    END;

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
