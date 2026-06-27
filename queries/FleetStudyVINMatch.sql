SELECT
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.ClientShortName AS FAClient,
  FleetStudyImportBase.UnitNumber AS ImportUnitNum,
  vw_SixKeys.Unitnum AS FA_UnitNum,
  vw_SixKeys.ObcUnitNumber,
  UMP_MilesFuelMPG.LastReport AS FA_LastReport,
  [FleetStudyImportBase].[UnitYear] & " " & [FleetStudyImportBase].[UnitMake] & " " & [FleetStudyImportBase].[UnitModel] & " " & [FleetStudyImportBase].[Type] & " " & [FleetStudyImportBase].[SubType] AS ImportDesc,
  FleetStudyImportBase.InServiceDate AS ImportISD,
  vw_SixKeys.UnitDescShortVer AS FA_Desc,
  vw_SixKeys.InServiceDate AS FA_ISD,
  IIf(
    CStr([UnitYear])<> CStr([MYear]),
    "Diff",
    "Match"
  ) AS YearCk,
  IIf(
    [vw_SixKeys].[Make] <> [UnitMake],
    "Diff", "Match"
  ) AS MakeCk,
  IIf(
    [UnitModel] <> [vw_SixKeys].[Model],
    "Diff", "Match"
  ) AS ModelCk,
  IIf(
    [FleetStudyImportBase].[InServiceDate] <> [vw_SixKeys].[InServiceDate],
    "Diff", "Match"
  ) AS ISDateCK,
  IIf(
    Trim(
      [FleetStudyImportBase].[SubType]
    )<> Trim([vw_SixKeys].[SubType]),
    "Diff",
    "Match"
  ) AS TypeCk,
  IIf(
    Nz(
      [UnitVinDetails].[model] & [UnitVinDetails].[make] & [UnitVinDetails].[manufacturer],
      ""
    )<> "",
    "Yes",
    "No"
  ) AS FA_VinCheck,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId
FROM
  FleetStudyImportBase
  INNER JOIN (
    (
      vw_SixKeys
      LEFT JOIN UMP_MilesFuelMPG ON vw_SixKeys.AssetID = UMP_MilesFuelMPG.UnitRef
    )
    LEFT JOIN UnitVinDetails ON vw_SixKeys.UnitID = UnitVinDetails.unitid
  ) ON FleetStudyImportBase.VIN = vw_SixKeys.VIN
ORDER BY
  vw_SixKeys.ClientShortName,
  [FleetStudyImportBase].[UnitYear] & " " & [FleetStudyImportBase].[UnitMake] & " " & [FleetStudyImportBase].[UnitModel] & " " & [FleetStudyImportBase].[Type] & " " & [FleetStudyImportBase].[SubType];
