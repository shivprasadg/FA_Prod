SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName AS FAClient,
  vw_SixKeys.Status AS FAUnitStatus,
  vw_SixKeys.AssetID,
  vw_SixKeys.Unitnum,
  vw_SixKeys.ObcUnitNumber,
  [UnitYear] & " " & [FleetStudyImport].[UnitMake] & " " & [FleetStudyImport].[UnitModel] & " " & [FleetStudyImport].[Type] & " " & [FleetStudyImport].[SubType] AS ImportDesc,
  vw_SixKeys.UnitDescShortVer AS FAUnitDesc,
  FleetStudyImport.VIN AS ImportVin,
  vw_SixKeys.VIN AS FA_Vin,
  [vw_SixKeys].[VIN] <> [FleetStudyImport].[VIN] AS VinDiff,
  IIf(
    Nz(
      [UnitVinDetails].[model] & [UnitVinDetails].[make] & [UnitVinDetails].[manufacturer],
      ""
    )<> "",
    "Yes",
    "No"
  ) AS FA_VinCheck,
  UMP_MilesFuelMPG.LastReport AS FA_LastReport,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId
FROM
  (
    (
      FleetStudyImport
      INNER JOIN vw_SixKeys ON FleetStudyImport.UnitNumber = vw_SixKeys.Unitnum
    )
    LEFT JOIN UMP_MilesFuelMPG ON vw_SixKeys.AssetID = UMP_MilesFuelMPG.UnitRef
  )
  LEFT JOIN UnitVinDetails ON vw_SixKeys.UnitID = UnitVinDetails.unitid
WHERE
  (
    (
      (vw_SixKeys.clientID)= [Forms]![FleetStudy_Importer]![clientID]
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.Status;
