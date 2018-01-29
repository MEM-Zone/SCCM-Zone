/*
*********************************************************************************************************
* Requires        | SCCM Hardware extentsion                                                            *
* ===================================================================================================== *
* Created by      |    Date    | Comments                                                               *
* _____________________________________________________________________________________________________ *
* Ioan Popovic    | 2018-01-29 | First version                                                          *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This SQL Query is used to get all collected OU's.
.DESCRIPTION
    This SQL Query is used to get all collected OU's by Company.
.NOTES
    Part of a report should not be run separately.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/

--DECLARE @UserSIDs VARCHAR(16);
--SELECT @UserSIDs = 'disabled'
--DECLARE @CollectionID VARCHAR(16);
--SELECT @CollectionID = 'WT10000A';

SELECT DISTINCT
  SystemOU.System_OU_Name0
FROM [v_RA_System_SystemOUName] AS [SystemOU]
WHERE 
    SystemOU.System_OU_Name0 LIKE '%%'
ORDER BY SystemOU.System_OU_Name0

/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
