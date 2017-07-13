DECLARE @AuthListLocalID AS int
SELECT @AuthListLocalID=CI_ID
FROM v_AuthListInfo
WHERE v_AuthListInfo.Ci_UniqueID=@AuthListID;


SELECT DISTINCT rs.NetBios_Name0 AS Name,
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
                    ELSE 'Other'
                END AS Osys,
                CASE
                    WHEN [cs].Roles0 LIKE '%Domain_Controller%' THEN 'DC'
                    WHEN [cs].Roles0 LIKE '%Domain_Controller%' THEN 'DC'
                    WHEN [cs].Roles0 LIKE '%Workstation%' THEN 'Workstation'
                    ELSE 'Other'
                END AS [Role],
                CASE
                    WHEN os.CSDVersion0 LIKE '%1%' THEN '1'
                    WHEN os.CSDVersion0 LIKE '%2%' THEN '2'
                    WHEN os.CSDVersion0 LIKE '%3%' THEN '3'
                    WHEN os.CSDVersion0 LIKE '%4%' THEN '4'
                    WHEN os.CSDVersion0 LIKE '%5%' THEN '5'
                END AS SP,
                ucsa.ResourceID,
                ui.BulletinID,
                ui.ArticleID,
                ui.Title,
                ui.Description,
                ui.DateRevised,
                CASE ui.Severity
                    WHEN 10 THEN 'Critical'
                    WHEN 8 THEN 'Important'
                    WHEN 6 THEN 'Moderate'
                    WHEN 2 THEN 'Low'
                    ELSE '(Unknown)'
                END AS [Severity]
FROM v_UpdateComplianceStatus ucsa
INNER JOIN v_CIRelation cir ON ucsa.CI_ID = cir.ToCIID
INNER JOIN v_UpdateInfo ui ON ucsa.CI_ID = ui.CI_ID
JOIN v_R_System rs ON ucsa.ResourceID = rs.ResourceID
LEFT JOIN dbo.v_GS_COMPUTER_SYSTEM CS ON rs.ResourceID = CS.ResourceID
JOIN v_GS_OPERATING_SYSTEM AS os ON ucsa.ResourceID = os.ResourceID
WHERE cir.RelationType=1
    AND ucsa.ResourceID IN
        (SELECT vc.ResourceID
         FROM v_FullCollectionMembership vc
         WHERE vc.CollectionID = @CollID)
    AND ucsa.Status = '2' --Required and [cir].[FromCIID] = @AuthListLocalID
