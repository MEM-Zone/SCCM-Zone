/*
.SYNOPSIS
    Gets the Update compliance in SCCM.
.DESCRIPTION
    Gets the Update compliance in SCCM by computer name and classification.
.NOTES
    Created by
        Ioan Popovici   2018-10-03
    Release notes
        --https://github.com/JhonnyTerminus/SCCMZone/blob/master/Reporting/Security/SEC%20BitLocker%20Compliance%20and%20Settings/CHANGELOG.md
    This query is part of a report should not be run separately.
.LINK
    https://SCCM-Zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCMZone
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/

/* Testing variables !! Need to be commented for Production !! */
DECLARE @UserSIDs       AS VARCHAR(16)  = 'Disabled';
DECLARE @CollectionID   AS VARCHAR(16)  = 'A01000B3';
DECLARE @Locale	        AS VARCHAR(2)   = '2';
DECLARE @LCID           AS INT          = dbo.fn_LShortNameToLCID (@Locale)

/* Get Update and Computer data */
SELECT DISTINCT
    Computer.Netbios_Name0 AS ComputerName
    , IPAddresses =
        REPLACE(
            (
                SELECT LTRIM(RTRIM(IP.IP_Addresses0)) AS [data()]
                FROM fn_rbac_RA_System_IPAddresses(@UserSIDs) AS IP
                WHERE IP.ResourceID = Computer.ResourceID
                    AND IP.IP_Addresses0 NOT LIKE 'fe%' -- Exclude IPv6
                FOR XML PATH('')
            ),
            ' ','; '
        )
    , UserName			= CONCAT(Computer.User_Domain0 + '\', Computer.User_Name0)                                          -- Add user domain to UserName
    , OperatingSystem   = CONCAT(OperatingSystem.Caption0, (REPLACE(OperatingSystem.CSDVersion0, 'Service Pack', ' SP')))   -- Add service pack to OperatingSystem
    , BuildNumber       = OperatingSystem.Version0
    , Version           = OSLocalizedNames.Value
    , Domain            = Computer.Full_Domain_Name0
    , Managed =
        CASE Computer.Client0
            WHEN 0 THEN 'No'
            WHEN 1 THEN 'Yes'
        END
    , ClientVersion     = Computer.Client_Version0
    , LastBootTime      = OperatingSystem.LastBootUpTime0
    , LastUpdateScan    = UpdateScan.LastScanTime
    , Status =
        CASE UCS.Status
            WHEN 0 THEN 'Detection state unknown'
            WHEN 1 THEN 'Update is not required'
            WHEN 2 THEN 'Update is required'
            WHEN 3 THEN 'Update is installed'
        END
    , Classification    = Category.CategoryInstanceName
    , Severity          = ISNULL(NULLIF(UI.SeverityName, ''), 'Unknown')
    , ArticleID         = UI.ArticleID
    , BulletinID        = NULLIF(UI.BulletinID, '')
    , DisplayName       = UI.DisplayName
    , DateRevised       = UI.DateRevised
    , Deployed =
        CASE UI.IsDeployed
            WHEN 0 THEN 'No'
            WHEN 1 THEN 'Yes'
        END
    , Enabled =
        CASE UI.IsEnabled
            WHEN 0 THEN 'No'
            WHEN 1 THEN 'Yes'
        END
FROM fn_rbac_R_System(@UserSIDs) AS Computer
    INNER JOIN fn_rbac_Update_ComplianceStatus(@UserSIDs) AS UCS ON Computer.ResourceID = UCS.ResourceID
        AND UCS.Status IN (0,2)     -- 0 Unknown, 2 Required, 3 Installed
    INNER JOIN fn_ListUpdateCIs(@Locale) AS UI ON UCS.CI_ID = UI.CI_ID
        AND UI.CIType_ID IN (1,8)   -- 1 Software Updates, 8 Software Update Bundle (v_CITypes)
        AND UI.IsExpired = 0        -- Update is not Expired
        AND UI.IsSuperseded = 0     -- Update is not Superseeded
    LEFT JOIN fn_rbac_GS_OPERATING_SYSTEM(@UserSIDs) OperatingSystem ON OperatingSystem.ResourceID = Computer.ResourceID
    LEFT JOIN fn_GetWindowsServicingStates() AS OSServicingStates ON OSServicingStates.Build = OperatingSystem.Version0
    LEFT JOIN fn_GetWindowsServicingLocalizedNames() AS OSLocalizedNames ON OSLocalizedNames.Name = OSServicingStates.Name
    LEFT JOIN fn_rbac_UpdateScanStatus(@UserSIDs) AS UpdateScan ON Computer.ResourceID = UpdateScan.ResourceID
    LEFT JOIN fn_rbac_CICategories_All(@UserSIDs) AS CICategories ON UI.CI_ID = CICategories.CI_ID
    RIGHT JOIN fn_rbac_ListUpdateCategoryInstances(@Locale, @UserSIDs) AS Category ON CICategories.CategoryInstanceID = Category.CategoryInstanceID
        AND Category.CategoryTypeName = 'UpdateClassification'  -- Get only the 'UpdateClasification' category
    LEFT JOIN fn_rbac_ClientCollectionMembers(@UserSIDs) AS CollectionMembers ON Computer.ResourceID = CollectionMembers.ResourceID
WHERE CollectionMembers.CollectionID = @CollectionID

/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/