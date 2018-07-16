/*
*********************************************************************************************************
* Created by Ioan Popovici   | Requires SQL                                                             *
* ===================================================================================================== *
* Modified by   |    Date    | Revision | Comments                                                      *
* _____________________________________________________________________________________________________ *
* Ioan Popovici | 2018-07-16 | v1.0     | First version                                                 *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This SQL Query is used to get installed software by collection and software name.
.DESCRIPTION
    This SQL Query is used to get installed software by collection and software name.
.NOTES
    Part of a report should not be run separately.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/
/* #region QueryBody */

/* Remove previous temporary table if exists */
IF OBJECT_ID (N'TempDB.DBO.#InstalledSoftware') IS NOT NULL
    BEGIN
        DROP TABLE #InstalledSoftware;
    END;

/* Get installed software */
SELECT DISTINCT
    SYS.Netbios_Name0 AS Computer,
    SW.Publisher0 AS Publisher,
    SW.DisplayName0 AS Software,
    SW.Version0 AS Version,
    SYS.Resource_Domain_OR_Workgr0 AS DomainOrWorkgroup,
    SYS.User_Name0 AS UserName,
    OS.Caption0 AS OperatingSystem
INTO #InstalledSoftware
FROM fn_rbac_Add_Remove_Programs(@UserSIDs) SW
    JOIN fn_rbac_R_System(@UserSIDs) SYS ON SW.ResourceID = SYS.ResourceID
    JOIN v_ClientCollectionMembers COL ON COL.ResourceID = SYS.ResourceID
    JOIN v_GS_OPERATING_SYSTEM OS ON OS.ResourceID = SYS.ResourceID
WHERE COL.CollectionID = @CollectionID
    AND SW.DisplayName0 LIKE '%'+@SoftwareName+'%'
ORDER BY SYS.Netbios_Name0,
    SW.Publisher0,
    SW.DisplayName0,
    SW.Version0;

/* Use NOT LIKE if needed */
IF @SoftwareNameNotLike != ''
BEGIN
    SELECT
        Computer,
        Publisher,
        Software,
        Version,
        DomainOrWorkgroup,
        UserName,
        OperatingSystem
    FROM #InstalledSoftware
        WHERE Software NOT LIKE '%'+@SoftwareNameNotLike+'%'
END;

/* Otherwise perform a normal select */
IF @SoftwareNameNotLike = ''
BEGIN
    SELECT
        Computer,
        Publisher,
        Software,
        Version,
        DomainOrWorkgroup,
        UserName,
        OperatingSystem
    FROM #InstalledSoftware
END;

/* Remove  temporary table */
DROP TABLE #InstalledSoftware;

/* #endregion */
/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/