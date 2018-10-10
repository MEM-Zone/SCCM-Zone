/*
.SYNOPSIS
    Gets the BitLocker compliance in SCCM.
.DESCRIPTION
    Gets the BitLocker compliance in SCCM, and displays verbose info.
.NOTES
    Created by
        Ioan Popovici   2018-10-03
    Release notes
        https://github.com/JhonnyTerminus/SCCMZone/blob/master/Reporting/Security/SEC%20BitLocker%20Compliance%20and%20Settings/CHANGELOG.md
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
DECLARE @UserSIDs       VARCHAR (16) = 'Disabled';
DECLARE @CollectionID   VARCHAR (16) = 'A01000B3';
DECLARE @ExcludeVirtualMachines   VARCHAR (3) = 'No';

/* Get BitLocker data */
SELECT
    Computer.Name0					AS ComputerName
	, Computer.Manufacturer0		AS Manufacturer,
    CASE
        WHEN Computer.Model0 LIKE '10AA%' THEN 'ThinkCentre M93p'
        WHEN Computer.Model0 LIKE '10AB%' THEN 'ThinkCentre M93p'
        WHEN Computer.Model0 LIKE '10AE%' THEN 'ThinkCentre M93z'
        WHEN Computer.Model0 LIKE '10FLS1TJ%' THEN 'ThinkCentre M900'
        WHEN Product.Version0 = 'Lenovo Product' THEN ('Unknown ' + Computer.Model0)
        WHEN Computer.Manufacturer0 = 'LENOVO' THEN Product.Version0
        ELSE Computer.Model0
    END AS Model
	, OperatingSystem.Caption0		AS OperatingSystem
	, OperatingSystem.CSDVersion0	AS ServicePack
    , OperatingSystem.Version0		AS BuildNumber
	, OSLocalizedNames.Value		AS Version
	, TPM.ManufacturerId0			AS ManufacturerID
	, TPM.ManufacturerVersion0		AS ManufacturerVersion
	, TPM.PhysicalPresenceVersionInfo0	AS PhisicalPresenceVersionInfo
	, TPM.SpecVersion0					AS SpecVersion

	, BitlockerPolicy =
	CAST (
        (
            SELECT (
                SELECT
                    ActiveDirectoryBackup0          AS ActiveDirectoryBackup
                    , ActiveDirectoryInfoToStore0   AS ActiveDirectoryInfoToStore
                    , CertificateOID0               AS CertificateOI
                    , DefaultRecoveryFolderPath0    AS DefaultRecoveryFolderPath
                    , DisableExternalDMAUnderLock0  AS DisableExternalDMAUnderLock
                    , DisallowStandardUserPINReset0 AS DisallowStandardUserPINReset
                    , EnableBDEWithNoTPM0           AS EnableBDEWithNoTPM
                    , EnableNonTPM0                 AS EnableNonTPM
                    , EncryptionMethod0             AS EnableNonTPM
                    , EncryptionMethodNoDiffuser0   AS EncryptionMethodNoDiffuser
                    , EncryptionMethodWithXtsFdv0   AS EncryptionMethodWithXtsFdv
                    , EncryptionMethodWithXtsOs0    AS EncryptionMethodWithXtsOs
                    , EncryptionMethodWithXtsRdv0   AS EncryptionMethodWithXtsRdv
                    , IdentificationField0          AS IdentificationField
                    , IdentificationFieldString0    AS IdentificationFieldString
                    , MinimumPIN0                   AS MinimumPIN
                    , MorBehavior0                  AS MorBehavior
                    , RecoveryKeyMessage0           AS RecoveryKeyMessage
                    , RecoveryKeyMessageSource0     AS RecoveryKeyMessageSource
                    , RecoveryKeyUrl0               AS RecoveryKeyUrl
                    , RequireActiveDirectoryBackup0 AS RequireActiveDirectoryBackup
                    , SecondaryIdentificationField0 AS SecondaryIdentificationField
                    , TPMAutoReseal0                AS TPMAutoReseal
                    , UseAdvancedStartup0           AS UseAdvancedStartup
                    , UseEnhancedPin0               AS UseEnhancedPin
                    , UsePartialEncryptionKey0      AS UsePartialEncryptionKey
                    , UsePIN0                       AS UsePIN
                    , UseRecoveryDrive0             AS UseRecoveryDrive
                    , UseRecoveryPassword0          AS UseRecoveryPassword
                    , UseTPM0                       AS UseTPM
                    , UseTPMKey0                    AS UseTPMKey
                    , UseTPMKeyPIN0                 AS UseTPMKeyPIN
                    , UseTPMPIN0                    AS UseTPMPIN
                FROM v_GS_CUSTOM_BITLOCKER_POLICY0
                WHERE ResourceID = BitLocker.ResourceID
                FOR XML PATH ('General'), TYPE
            ),
            (
                SELECT
                    OSActiveDirectoryBackup0        AS OSActiveDirectoryBackup
                    , OSActiveDirectoryInfoToStore0 AS OSActiveDirectoryInfoToStore
                    , OSAllowedHardwareEncryptionA0 AS OSAllowedHardwareEncryptionAlgorithms
                    , OSAllowSecureBootForIntegrit0 AS OSAllowSecureBootForIntegrity
                    , OSAllowSoftwareEncryptionFai0 AS OSAllowSoftwareEncryptionFailover
                    , OSBcdAdditionalExcludedSetti0 AS OSBcdAdditionalExcludedSettings
                    , OSBcdAdditionalSecurityCriti0 AS OSBcdAdditionalSecurityCriticalSettings
                    , OSEnablePrebootInputProtecto0 AS OSEnablePrebootInputProtectorsOnSlates
                    , OSEnablePreBootPinExceptionO0 AS OSEnablePreBootPinExceptionOnDECapableDevice
                    , OSEncryptionType0             AS OSEncryptionType
                    , OSHardwareEncryption0         AS OSHardwareEncryption
                    , OSHideRecoveryPage0           AS OSHideRecoveryPage
                    , OSManageDRA0                  AS OSManageDRA
                    , OSManageNKP0                  AS OSManageNKP
                    , OSPassphrase0                 AS OSPassphrase
                    , OSPassphraseASCIIOnly0        AS OSPassphraseASCIIOnly
                    , OSPassphraseComplexity0       AS OSPassphraseComplexity
                    , OSPassphraseLength0           AS OSPassphraseLength
                    , OSRecovery0                   AS OSRecovery
                    , OSRecoveryKey0                AS OSRecoveryKey
                    , OSRecoveryPassword0           AS OSRecoveryPassword
                    , OSRequireActiveDirectoryBack0 AS OSRequireActiveDirectoryBackup
                    , OSRestrictHardwareEncryption0 AS  OSRestrictHardwareEncryptionAlgorithms
                    , OSUseEnhancedBcdProfile0      AS OSUseEnhancedBcdProfile
                FROM v_GS_CUSTOM_BITLOCKER_POLICY0
                WHERE ResourceID = BitLocker.ResourceID
                FOR XML PATH ('OSDrives'), TYPE
            ),
            (
                SELECT
                    FDVActiveDirectoryBackup0       AS FDVActiveDirectoryBackup
                    , FDVActiveDirectoryInfoToStor0 AS FDVActiveDirectoryInfoToStore
                    , FDVAllowedHardwareEncryption0 AS FDVAllowedHardwareEncryptionAlgorithms
                    , FDVAllowSoftwareEncryptionFa0 AS FDVAllowSoftwareEncryptionFailover
                    , FDVAllowUserCert0             AS FDVAllowUserCert
                    , FDVDiscoveryVolumeType0       AS FDVDiscoveryVolumeType
                    , FDVEncryptionType0            AS FDVEncryptionType
                    , FDVEnforcePassphrase0         AS FDVEnforcePassphrase
                    , FDVEnforceUserCert0           AS FDVEnforceUserCert
                    , FDVHardwareEncryption0        AS FDVHardwareEncryption
                    , FDVHideRecoveryPage0          AS FDVHideRecoveryPage
                    , FDVManageDRA0                 AS FDVManageDRA
                    , FDVNoBitLockerToGoReader0     AS FDVNoBitLockerToGoReader
                    , FDVPassphrase0                AS FDVPassphrase
                    , FDVPassphraseComplexity0      AS FDVPassphraseComplexity
                    , FDVPassphraseLength0          AS FDVPassphraseLength
                    , FDVRecovery0                  AS FDVRecovery
                    , FDVRecoveryKey0               AS FDVRecoveryKey
                    , FDVRecoveryPassword0          AS FDVRecoveryPassword
                    , FDVRequireActiveDirectoryBac0 AS FDVRequireActiveDirectoryBackup
                    , FDVRestrictHardwareEncryptio0 AS FDVRestrictHardwareEncryptionAlgorithms
                FROM v_GS_CUSTOM_BITLOCKER_POLICY0
                WHERE ResourceID = BitLocker.ResourceID
                FOR XML PATH ('FixedDrives'), TYPE
            ),
            (
                SELECT
                    RDVActiveDirectoryBackup0       AS RDVActiveDirectoryBackup0
                    , RDVActiveDirectoryInfoToStor0 AS RDVActiveDirectoryInfoToStore
                    , RDVAllowBDE0                  AS RDVAllowBDE
                    , RDVAllowedHardwareEncryption0 AS RDVAllowedHardwareEncryptionAlgorithms
                    , RDVAllowSoftwareEncryptionFa0 AS RDVAllowSoftwareEncryptionFailover
                    , RDVAllowUserCert0             AS RDVAllowUserCert
                    , RDVConfigureBDE0              AS RDVConfigureBDE
                    , RDVDenyCrossOrg0              AS RDVDenyCrossOrg
                    , RDVDisableBDE0                AS RDVDisableBDE
                    , RDVDiscoveryVolumeType0       AS RDVDiscoveryVolumeType
                    , RDVEncryptionType0            AS RDVEncryptionType
                    , RDVEnforcePassphrase0         AS RDVEnforcePassphrase
                    , RDVEnforceUserCert0           AS RDVEnforceUserCert
                    , RDVHardwareEncryption0        AS RDVHardwareEncryption
                    , RDVHideRecoveryPage0          AS RDVHideRecoveryPage
                    , RDVManageDRA0                 AS RDVManageDRA
                    , RDVNoBitLockerToGoReader0     AS RDVNoBitLockerToGoReader
                    , RDVPassphrase0                AS RDVPassphrase
                    , RDVPassphraseComplexity0      AS RDVPassphraseComplexity
                    , RDVPassphraseLength0          AS RDVPassphraseLength
                    , RDVRecovery0                  AS RDVRecovery
                    , RDVRecoveryKey0               AS RDVRecoveryKey
                    , RDVRecoveryPassword0          AS RDVRecoveryPassword
                    , RDVRequireActiveDirectoryBac0 AS RDVRequireActiveDirectoryBackup
                    , RDVRestrictHardwareEncryptio0 AS RDVRestrictHardwareEncryptionAlgorithms
                FROM v_GS_CUSTOM_BITLOCKER_POLICY0
                WHERE ResourceID = BitLocker.ResourceID
                FOR XML PATH ('RemovableDrives'), TYPE
            )
            FOR XML PATH (''), ROOT ('BitLocker')
        ) AS XML
    )
    , IsVolumeInitializedForProtection =
    (
		CASE BitLocker.IsVolumeInitializedForProtec0
			WHEN 0 THEN 'No'
			WHEN 1 THEN 'Yes'
        END
	)
    , BitLocker.DriveLetter0 AS DriveLetter
    , ProtectionStatus =
    (
		CASE BitLocker.ProtectionStatus0
			WHEN 0 THEN 'OFF'
			WHEN 1 THEN 'ON'
			WHEN 2 THEN 'UNKNOWN'
        END
	)
    , ConversionStatus =
    (
		CASE BitLocker.ConversionStatus0
			WHEN 0 THEN 'FullyDecrypted'
			WHEN 1 THEN 'FullyEncrypted'
			WHEN 2 THEN 'EncryptionInProgress'
			WHEN 3 THEN 'DecryptionInProgress'
			WHEN 4 THEN 'EncryptionPaused'
			WHEN 5 THEN 'DecryptionPaused'
        END
	)
    , EncryptionMethod =
    (
		CASE BitLocker.EncryptionMethod0
			WHEN 0 THEN 'None'
			WHEN 1 THEN 'AES_128_WITH_DIFFUSER'
			WHEN 2 THEN 'AES_256_WITH_DIFFUSER'
			WHEN 3 THEN 'AES_128'
			WHEN 4 THEN 'AES_256'
			WHEN 5 THEN 'HARDWARE_ENCRYPTION'
			WHEN 6 THEN 'XTS_AES_128'
			WHEN 7 THEN 'XTS_AES_256'
			WHEN -1 THEN 'UNKNOWN'
        END
	)
    , VolumeType =
    (
		CASE BitLocker.VolumeType0
			WHEN 0 THEN 'OSVolume'
			WHEN 1 THEN 'FixedDataVolume'
			WHEN 2 THEN 'PortableDataVolume'
			WHEN 3 THEN 'VirtualDataVolume'
        END
    )
    , DeviceID = (
        SELECT SUBSTRING (
            DeviceID0,
            CHARINDEX ('{', DeviceID0) + LEN ('{'),
            CHARINDEX ('}', DeviceID0) - CHARINDEX ('{', DeviceID0) - LEN ('{')
        )
	)
FROM
    dbo.fn_rbac_GS_COMPUTER_SYSTEM (@UserSIDs) AS Computer
    LEFT JOIN dbo.v_GS_CUSTOM_ENCRYPTABLE_VOLUME_EXT0 AS BitLocker ON BitLocker.ResourceID = Computer.ResourceID
    LEFT JOIN dbo.v_GS_CUSTOM_BITLOCKER_POLICY0 AS BitLockerPolicy ON BitLockerPolicy.ResourceID = Computer.ResourceID
    LEFT JOIN dbo.v_GS_OPERATING_SYSTEM OperatingSystem ON OperatingSystem.ResourceID = Computer.ResourceID
    LEFT JOIN dbo.v_GS_COMPUTER_SYSTEM_PRODUCT AS Product ON Product.ResourceID = Computer.ResourceID
    LEFT JOIN fn_rbac_GS_TPM (@UserSIDs) Tpm on Tpm.ResourceID = Computer.ResourceID
    LEFT JOIN dbo.v_ClientCollectionMembers AS Collection ON Collection.ResourceID = Computer.ResourceID
    Left JOIN dbo.vSMS_WindowsServicingStates AS OSServicingStates ON OSServicingStates.Build = OperatingSystem.Version0
    Left JOIN vSMS_WindowsServicingLocalizedNames AS OSLocalizedNames ON OSLocalizedNames.Name = OSServicingStates.Name
WHERE
	Collection.CollectionID = @CollectionID
    AND
    Computer.Model0 NOT LIKE (
            CASE @ExcludeVirtualMachines
                WHEN 'YES' THEN '%Virtual%'
                ELSE ''
            END
    )
GROUP BY
	Computer.Name0
	,Computer.Manufacturer0
	,Computer.Model0
	,Product.Version0
	,OperatingSystem.Caption0
	,OperatingSystem.Version0
	,OperatingSystem.CSDVersion0
	,OSLocalizedNames.Value
	,TPM.ManufacturerId0
	,TPM.ManufacturerVersion0
	,TPM.PhysicalPresenceVersionInfo0
	,TPM.SpecVersion0
	,BitLocker.ResourceID
	,BitLocker.IsVolumeInitializedForProtec0
	,BitLocker.DriveLetter0
	,BitLocker.ProtectionStatus0
	,BitLocker.ConversionStatus0
	,BitLocker.EncryptionMethod0
	,BitLocker.VolumeType0
	,BitLocker.DeviceID0


/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/