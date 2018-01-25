/*
*********************************************************************************************************
* Requires        | SCCM Hardware extentsion                                                            *
* ===================================================================================================== *
* Created by      |    Date    | Comments                                                               *
* _____________________________________________________________________________________________________ *
* Octavian Cordos | 2016-01-15 | First version                                                          *
* Ioan Popovici   |            |                                                                        *
* ===================================================================================================== *
*                                                                                                       *
*********************************************************************************************************

.SYNOPSIS
    This SQL Query is used to get SQL Service Pack and Cumulative Update version information.
.DESCRIPTION
    This SQL Query is used to get SQL Service Pack and Cumulative Update version information.
.NOTES
    Part of a report should not be run separately.
.LINK
    https://sccm-zone.com
    https://github.com/JhonnyTerminus/SCCM
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/

--DECLARE @CollectionID VARCHAR(16);
--SELECT @CollectionID = 'WT10000A';

WITH temp (
        [ResourceID],
        [ComputerName],
        --[Company],
        [SQL TYPE],
        [SQL Service Pack],
        [SQL Version],
        [SQL CU Version],
        [Version]
    ) AS (
        SELECT DISTINCT
            [ResourceID],
            [ComputerName],
            --[Company],
            [SQL TYPE],
            [SQL Service Pack],
            [SQL Version],
            [SQL CU Version],
            [Version]
        FROM (

--region SQL 2017
            SELECT
                [vrs].[resourceID] AS [ResourceID],
                [VRS].[Netbios_name0] [ComputerName],
                --ISNULL([vrs].[company0], '<Unknown>') AS  'Company',
                --vrs.company0 as 'Company',
                MAX (
                    CASE [sql2017].[PropertyName0]
                        WHEN 'SKUName' THEN [sql2017].[PropertySTRValue0]
                    END
                ) AS [SQL TYPE],
                MAX (
                    CASE [sql2017].[PropertyName0]
                        WHEN 'SPLEVEL' THEN [sql2017].[PropertyNUMValue0]
                    END
                ) AS [SQL Service Pack],
                MAX (
                    CASE [sql2017].[PropertyName0]
                        WHEN 'VERSION' THEN [sql2017].[PropertySTRValue0]
                    END
                ) AS [SQL Version],
                MAX (
                    CASE [sql2017].[PropertyName0]
                        WHEN 'FILEVERSION' THEN [sql2017].[PropertySTRValue0]
                    END
                ) AS [SQL CU Version],
                MAX (
                    CASE [sql2017].[PropertyName0]
                        WHEN 'FILEVERSION' THEN
                            CASE LEFT ([sql2017].[PropertySTRValue0], 4)
                                WHEN '2017' THEN '2017'
                                WHEN '2016' THEN '2016'
                                WHEN '2014' THEN '2014'
                                WHEN '2011' THEN '2012'
                                WHEN '2009' THEN '2008 R2'
                                WHEN '2007' THEN '2008'
                                WHEN '2005' THEN '2005'
                                WHEN '2000' THEN '2000'
                                ELSE '2016'
                            END
                    END
                ) AS [Version]
            FROM [dbo].[v_R_System] [VRS]
                LEFT JOIN [v_GS_SQL_2017_Property0] [sql2017] ON [sql2017].[ResourceID] = [VRS].[ResourceID]
                LEFT OUTER JOIN [v_ClientCollectionMembers] [c] ON [c].[ResourceID] = [vrs].[ResourceID]
            WHERE [sql2017].[PropertyName0] IN ('SKUNAME', 'SPLevel', 'version', 'fileversion')
                AND [c].[CollectionID] = @CollectionID
                AND ISNULL([sql2017].[ServiceName0], 0) NOT LIKE '%EXPRESS%'
            GROUP BY
                [VRS].[Netbios_name0],
                [sql2017].[ServiceName0],
                --[vrs].[company0],
                [vrs].[resourceID]
            UNION ALL
--endregion

--region SQL 2016
            SELECT
                [vrs].[resourceID] [ResourceID],
                [VRS].[Netbios_name0] [ComputerName],
                --ISNULL([vrs].[company0], '<Unknown>') AS 'Company',
                --vrs.company0 as 'Company',
                MAX (
                    CASE [sql2016].[PropertyName0]
                        WHEN 'SKUName' THEN [sql2016].[PropertySTRValue0]
                    END
                ) AS [SQL TYPE],
                MAX (CASE [sql2016].[PropertyName0]
                        WHEN 'SPLEVEL' THEN [sql2016].[PropertyNUMValue0]
                    END
                ) AS [SQL Service Pack],
                MAX (
                    CASE [sql2016].[PropertyName0]
                        WHEN 'VERSION' THEN [sql2016].[PropertySTRValue0]
                    END
                ) AS [SQL Version],
                MAX (
                    CASE [sql2016].[PropertyName0]
                        WHEN 'FILEVERSION' THEN [sql2016].[PropertySTRValue0]
                    END
                ) AS [SQL CU Version],
                MAX (
                    CASE [sql2016].[PropertyName0]
                        WHEN 'FILEVERSION' THEN
                            CASE LEFT ([sql2016].[PropertySTRValue0], 4)
                                WHEN '2017' THEN '2017'
                                WHEN '2016' THEN '2016'
                                WHEN '2014' THEN '2014'
                                WHEN '2011' THEN '2012'
                                WHEN '2009' THEN '2008 R2'
                                WHEN '2007' THEN '2008'
                                WHEN '2005' THEN '2005'
                                WHEN '2000' THEN '2000'
                                ELSE '2016'
                            END
                    END
                ) AS [Version]
            FROM [dbo].[v_R_System] [VRS]
                LEFT JOIN [v_GS_SQL_2016_Property0] [sql2016] ON [sql2016].[ResourceID] = [VRS].[ResourceID]
                LEFT OUTER JOIN [v_ClientCollectionMembers] [c] ON [c].[ResourceID] = [vrs].[ResourceID]
            WHERE [sql2016].[PropertyName0] IN ('SKUNAME', 'SPLevel', 'version', 'fileversion')
                AND [c].[CollectionID] = @CollectionID
                AND ISNULL([sql2016].[ServiceName0], 0) NOT LIKE '%EXPRESS%'
            GROUP BY [VRS].[Netbios_name0],
                [sql2016].[ServiceName0],
                --[vrs].[company0],
                [vrs].[resourceID]
--endregion

--region SQL 2014
            UNION ALL
            SELECT
                [vrs].[resourceID] [ResourceID],
                [VRS].[Netbios_name0] [ComputerName],
                --ISNULL([vrs].[company0], '<Unknown>') AS 'Company',
                --vrs.company0 as 'Company',
                MAX (
                    CASE [sql2014].[PropertyName0]
                        WHEN 'SKUName' THEN [sql2014].[PropertySTRValue0]
                    END
                ) AS [SQL TYPE],
                MAX (
                    CASE [sql2014].[PropertyName0]
                        WHEN 'SPLEVEL' THEN [sql2014].[PropertyNUMValue0]
                    END
                ) AS [SQL Service Pack],
                MAX (CASE [sql2014].[PropertyName0]
                        WHEN 'VERSION' THEN [sql2014].[PropertySTRValue0]
                    END
                ) AS [SQL Version],
                MAX (
                    CASE [sql2014].[PropertyName0]
                        WHEN 'FILEVERSION' THEN [sql2014].[PropertySTRValue0]
                    END
                ) AS [SQL CU Version],
                MAX (
                    CASE [sql2014].[PropertyName0]
                        WHEN 'FILEVERSION' THEN
                            CASE LEFT([sql2014].[PropertySTRValue0], 4)
                                WHEN '2017' THEN '2017'
                                WHEN '2016' THEN '2016'
                                WHEN '2014' THEN '2014'
                                WHEN '2011' THEN '2012'
                                WHEN '2009' THEN '2008 R2'
                                WHEN '2007' THEN '2008'
                                WHEN '2005' THEN '2005'
                                WHEN '2000' THEN '2000'
                                ELSE '2014'
                            END
                    END
                ) AS [Version]
            FROM [v_R_System] [VRS]
                LEFT JOIN [v_GS_SQL_2014_Property0] [sql2014] ON [sql2014].[ResourceID] = [VRS].[ResourceID]
                LEFT OUTER JOIN [v_ClientCollectionMembers] [c] ON [c].[ResourceID] = [vrs].[ResourceID]
            WHERE [sql2014].[PropertyName0] IN ('SKUNAME', 'SPLevel', 'version', 'fileversion')
                AND [c].[CollectionID] = @CollectionID
                AND ISNULL([sql2014].[ServiceName0], 0) NOT LIKE '%EXPRESS%'
            GROUP BY
                [VRS].[Netbios_name0],
                [sql2014].[ServiceName0],
                --[vrs].[company0],
                [vrs].[resourceID]
--endregion

--region SQL 2012
            UNION ALL
            SELECT
                [vrs].[resourceID] [ResourceID],
                [VRS].[Netbios_name0] [ComputerName],
                --ISNULL([vrs].[company0], '<Unknown>') AS 'Company',
                MAX (
                    CASE [sql2012].[PropertyName0]
                        WHEN 'SKUName' THEN [sql2012].[PropertySTRValue0]
                    END
                ) AS [SQL TYPE],
                MAX (
                    CASE [sql2012].[PropertyName0]
                        WHEN 'SPLEVEL' THEN [sql2012].[PropertyNUMValue0]
                    END
                ) AS [SQL Service Pack],
                MAX (
                    CASE [sql2012].[PropertyName0]
                        WHEN 'VERSION' THEN [sql2012].[PropertySTRValue0]
                    END
                ) AS [SQL Version],
                MAX (CASE [sql2012].[PropertyName0]
                        WHEN 'FILEVERSION' THEN [sql2012].[PropertySTRValue0]
                    END
                ) AS [SQL CU Version],
                MAX (
                    CASE [sql2012].[PropertyName0]
                        WHEN 'FILEVERSION' THEN
                            CASE LEFT([sql2012].[PropertySTRValue0], 4)
                                WHEN '2017' THEN '2017'
                                WHEN '2016' THEN '2016'
                                WHEN '2014' THEN '2014'
                                WHEN '2011' THEN '2012'
                                WHEN '2009' THEN '2008 R2'
                                WHEN '2007' THEN '2008'
                                WHEN '2005' THEN '2005'
                                WHEN '2000' THEN '2000'
                                ELSE '2012'
                            END
                    END
                ) AS [Version]
            FROM [V_R_System] [VRS]
                LEFT JOIN [v_GS_SQL_2012_Property0] [sql2012] ON [sql2012].[ResourceID] = [VRS].[ResourceID]
                LEFT OUTER JOIN [v_ClientCollectionMembers] [c] ON [c].[ResourceID] = [vrs].[ResourceID]
            WHERE [sql2012].[PropertyName0] IN('SKUNAME', 'SPLevel', 'version', 'fileversion')
                AND [c].[CollectionID] = @CollectionID
                AND ISNULL([ServiceName0], 0) NOT LIKE '%EXPRESS%'
            GROUP BY
                [VRS].[Netbios_name0],
                [sql2012].[ServiceName0],
                --[vrs].[company0],
                [vrs].[resourceID]
--endregion

--region SQL 2008
            UNION ALL
            SELECT
                [vrs].[resourceID] [ResourceID],
                [VRS].[Netbios_name0] [ComputerName],
                --ISNULL([vrs].[company0], '<Unknown>') AS 'Company',
                MAX (
                    CASE [sql2008].[PropertyName0]
                        WHEN 'SKUName' THEN [sql2008].[PropertySTRValue0]
                    END
                ) AS [SQL TYPE],
                MAX (
                    CASE [sql2008].[PropertyName0]
                        WHEN 'SPLEVEL' THEN [sql2008].[PropertyNUMValue0]
                    END
                ) AS [SQL Service Pack],
                MAX (
                    CASE [sql2008].[PropertyName0]
                        WHEN 'VERSION' THEN [sql2008].[PropertySTRValue0]
                    END
                ) AS [SQL Version],
                MAX (
                    CASE [sql2008].[PropertyName0]
                        WHEN 'FILEVERSION' THEN [sql2008].[PropertySTRValue0]
                    END
                ) AS [SQL CU Version],
                MAX (
                    CASE [sql2008].[PropertyName0]
                        WHEN 'FILEVERSION' THEN
                            CASE LEFT([sql2008].[PropertySTRValue0], 4)
                                WHEN '2017' THEN '2017'
                                WHEN '2016' THEN '2016'
                                WHEN '2014' THEN '2014'
                                WHEN '2011' THEN '2012'
                                WHEN '2009' THEN '2008 R2'
                                WHEN '2007' THEN '2008'
                                WHEN '2005' THEN '2005'
                                WHEN '2000' THEN '2000'
                                ELSE '2008'
                            END
                    END
                ) AS [Version]
            FROM [V_R_System] [VRS]
                LEFT JOIN [v_GS_SQL_2008_Property0] [sql2008] ON [sql2008].[ResourceID] = [VRS].[ResourceID]
                LEFT OUTER JOIN [v_ClientCollectionMembers] [c] ON [c].[ResourceID] = [vrs].[ResourceID]
            WHERE [sql2008].[PropertyName0] IN ('SKUNAME', 'SPLevel', 'version', 'fileversion')
                AND [c].[CollectionID] = @CollectionID
                AND ISNULL([sql2008].[ServiceName0], 0) NOT LIKE '%EXPRESS%'
                AND ISNULL([sql2008].[ServiceName0], 0) NOT LIKE 'SQLBrowser'
            GROUP BY
                [VRS].[Netbios_name0],
                [sql2008].[ServiceName0],
                --[vrs].[company0],
                [vrs].[resourceID]
--endregion

--region SQL Legacy
            UNION ALL
            SELECT [vrs].[resourceID] [ResourceID],
                [VRS].[Netbios_name0] [ComputerName],
                --ISNULL([vrs].[company0], '<Unknown>') AS 'Company',
                MAX (
                    CASE [sqlLgcy].[PropertyName0]
                        WHEN 'SKUName' THEN [sqlLgcy].[PropertySTRValue0]
                    END
                ) AS [SQL TYPE],
                MAX (
                    CASE [sqlLgcy].[PropertyName0]
                        WHEN 'SPLEVEL' THEN [sqlLgcy].[PropertyNUMValue0]
                    END
                ) AS [SQL Service Pack],
                MAX (
                    CASE [sqlLgcy].[PropertyName0]
                        WHEN 'VERSION' THEN [sqlLgcy].[PropertySTRValue0]
                    END
                ) AS [SQL Version],
                MAX (
                    CASE [sqlLgcy].[PropertyName0]
                        WHEN 'FILEVERSION' THEN [sqlLgcy].[PropertySTRValue0]
                    END
                ) AS [SQL CU Version],
                MAX (
                    CASE [sqllgcy].[PropertyName0]
                        WHEN 'FILEVERSION' THEN
                        CASE LEFT ([sqllgcy].[PropertySTRValue0], 4)
                            WHEN '2017' THEN '2017'
                            WHEN '2016' THEN '2016'
                            WHEN '2014' THEN '2014'
                            WHEN '2011' THEN '2012'
                            WHEN '2009' THEN '2008 R2'
                            WHEN '2007' THEN '2008'
                            WHEN '2005' THEN '2005'
                            WHEN '2000' THEN '2000'
                            ELSE '2005'
                        END
                    END
                ) AS [Version]
            FROM [V_R_System] [VRS]
                LEFT JOIN [v_GS_SQL_Legacy_Property0] [sqlLgcy] ON [sqlLgcy].[ResourceID] = [VRS].[ResourceID]
                LEFT OUTER JOIN [v_ClientCollectionMembers] [c] ON [c].[ResourceID] = [vrs].[ResourceID]
            WHERE [sqlLgcy].[PropertyName0] IN('SKUNAME', 'SPLevel', 'version', 'fileversion')
                AND [c].[CollectionID] = @CollectionID
                AND ISNULL([sqlLgcy].[ServiceName0], 0) NOT LIKE '%EXPRESS%'
                AND ISNULL([sqlLgcy].[ServiceName0], 0) NOT LIKE 'SQLBrowser'
            GROUP BY
                [VRS].[Netbios_Name0],
                [sqlLgcy].[ServiceName0],
                --[vrs].[company0],
                [vrs].[resourceID]
        ) AS [SQLInv]
--endregion

        WHERE [SQL TYPE] NOT LIKE 'Express%'
            AND [SQL TYPE] NOT LIKE 'Windows Internal Database%'
            AND SUBSTRING ( [SQL Version], 1, 2) != SUBSTRING ([SQL CU Version], 1, 2) )
        SELECT DISTINCT
            --[Company],
            [ComputerName],
            [SQL Type],
            [SQL Service Pack] AS [Service Pack],
            [SQL Version] AS Version,
            [SQL CU Version] AS [CU Version],
            [version] AS Release,
            (
                CASE
                    WHEN [SQL Type] LIKE '%workgroup%' THEN 'Workgroup Edition'
                    WHEN [SQL Type] LIKE '%develop%' THEN 'Developer Edition'
                    WHEN [SQL Type] LIKE '%standard%' THEN 'Standard Edition'
                    WHEN [SQL Type] LIKE '%enterprise%' THEN 'Enterprise Edition'
                    ELSE version
                END
            ) AS 'Edition',
            (
                CASE
                    WHEN [SQL Type] LIKE '%64%' THEN 'x64'
                    ELSE 'x32'
                END
            ) AS 'Bitness'
        FROM temp
        ORDER BY
            --temp.[Company],
            Release,
            Edition,
            Bitness,
            Version,
            ComputerName;

/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/
