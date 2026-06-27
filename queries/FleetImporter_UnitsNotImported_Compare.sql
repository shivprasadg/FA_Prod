SELECT
  vw_SixKeys.ClientShortName,
  FleetImporter_UnitsNotImported.UnitNumber,
  FleetImporter_UnitsNotImported.OBCUnitNumber,
  FleetImporter_UnitsNotImported.VIN,
  CInt(
    [FleetImporter_UnitsNotImported].[UnitYear]
  ) AS UnitYear,
  FleetImporter_UnitsNotImported.UnitMake,
  FleetImporter_UnitsNotImported.UnitModel,
  FleetImporter_UnitsNotImported.Type,
  FleetImporter_UnitsNotImported.SubType,
  vw_SixKeys.AssetID,
  vw_SixKeys.ObcUnitNumber AS FA_OBCNum,
  vw_SixKeys.Status AS FA_Status,
  dbo_vw_LastMilesReportedAll.LastReport AS OBC_LastReport,
  vw_SixKeys.MYear AS FA_ModelYear,
  vw_SixKeys.Make AS FA_Make,
  vw_SixKeys.Model AS FA_Model,
  vw_SixKeys.SubType AS FA_SubType,
  vw_SixKeys.Type AS FA_Type,
  vw_SixKeys.VIN AS FA_VIN
FROM
  (
    FleetImporter_UnitsNotImported
    INNER JOIN vw_SixKeys ON FleetImporter_UnitsNotImported.UnitNumber = vw_SixKeys.Unitnum
  )
  LEFT JOIN dbo_vw_LastMilesReportedAll ON vw_SixKeys.AssetID = dbo_vw_LastMilesReportedAll.UnitID
WHERE
  (
    (
      (vw_SixKeys.clientID)= [Forms]![FleetStudy_Importer]![clientID]
    )
  );
