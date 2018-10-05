/*
.SYNOPSIS
    Gets the BitLocker compliance in SCCM.
.DESCRIPTION
    Gets the BitLocker compliance in SCCM, and displays verbose info.
.NOTES
    Created by
        Ioan Popovici   2018-10-05
    Release notes
        https://github.com/JhonnyTerminus/SCCMZone/blob/master/Reporting/Security/SEC%20Bitlocker%20Compliance%20and%20Settings/CHANGELOG.md
    This query is part of a report should not be run separately.
.LINK
    https://SCCM-Zone.com
.LINK
    https://github.com/JhonnyTerminus/SCCMZone
*/

/*##=============================================*/
/*## QUERY BODY                                  */
/*##=============================================*/

/* Get BitLocker data */
SELECT [ResourceID]
    ,[DeviceID0]
    ,[IsVolumeInitializedForProtec0]
    ,DriveLetter0 AS DriveLetter
    ,ProtectionStatus =
	(
		CASE ProtectionStatus0
			WHEN 0 THEN 'Protection OFF'
			WHEN 1 THEN 'Protection ON'
			WHEN 2 THEN 'Protection UNKNOWN'
		END
	)
    ,ConversionStatus =
	(
		CASE ConversionStatus0
			WHEN 0 THEN 'FullyDecrypted'
			WHEN 1 THEN 'FullyEncrypted'
			WHEN 2 THEN 'EncryptionInProgress'
			WHEN 3 THEN 'DecryptionInProgress'
			WHEN 4 THEN 'EncryptionPaused'
			WHEN 5 THEN 'DecryptionPaused'
		END
	)
    ,EncryptionMethod =
	(
		CASE EncryptionMethod0
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
    ,VolumeType =
    (
		CASE VolumeType0
			WHEN 0 THEN 'OSVolume'
			WHEN 1 THEN 'FixedDataVolume'
			WHEN 2 THEN 'PortableDataVolume'
			WHEN 3 THEN 'VirtualDataVolume'
		END
    )
FROM
    v_GS_ENCRYPTABLE_VOLUME_EXT

/*##=============================================*/
/*## END QUERY BODY                              */
/*##=============================================*/