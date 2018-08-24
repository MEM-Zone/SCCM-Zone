/*
*********************************************************************************************************
* Requires          | SQL      																			*
* ===================================================================================================== *
* Modified by       |    Date    | Revision | Comments                                                  *
* _____________________________________________________________________________________________________ *
* Ioan Popovici     | 2018-08-21 | First version    						        					*
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This SQL Query is used to get the Maintenance Window info of a Computer Collection.
.DESCRIPTION
    This SQL Query is used to get the Maintenance Window info of a Computer Collection including Start Time and Duration.
.NOTES
    Part of a report should not be run separately.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/

SELECT
	Collection.Name AS Collection,
	ServiceWindow.Name,
	ServiceWindow.Description,
	CASE ServiceWindow.ServiceWindowType
		WHEN 1 THEN 'All Deployments'
		WHEN 2 THEN 'Programs'
		WHEN 3 THEN 'Reboot Required'
		WHEN 4 THEN 'Software Updates'
		WHEN 5 THEN 'Task Sequences'
		WHEN 6 THEN 'User Defined'
	END AS Type,
	ServiceWindow.StartTime, 
	ServiceWindow.Duration,
	CASE ServiceWindow.IsEnabled
		WHEN 1 THEN 'Yes'
		ELSE 'No'
	END AS Enabled
FROM dbo.fn_rbac_ServiceWindow(@UserSIDs) AS ServiceWindow
	JOIN v_Collection AS Collection ON Collection.CollectionID = ServiceWindow.CollectionID
ORDER BY Name

/* Use NOT LIKE if needed */
IF @SoftwareNameNotLike != ''
BEGIN
    SELECT
        Computer,
        Manufacturer,
        ComputerType,
        SerialNumber,
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
        Manufacturer,
        ComputerType,
        SerialNumber,
        Publisher,
        Software,
        Version,
        DomainOrWorkgroup,
        UserName,
        OperatingSystem
    FROM #InstalledSoftware
END;

/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/