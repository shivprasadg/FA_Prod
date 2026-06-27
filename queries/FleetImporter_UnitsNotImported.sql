SELECT
  FleetStudyImportBase.UnitNumber,
  FleetStudyImportBase.OBCUnitNumber,
  FleetStudyImportBase.VIN,
  FleetStudyImportBase.UnitYear,
  FleetStudyImportBase.UnitMake,
  FleetStudyImportBase.UnitModel,
  FleetStudyImportBase.Type,
  FleetStudyImportBase.SubType,
  FleetStudyImportBase.InServiceDate,
  FleetStudyImportBase.Axle,
  FleetStudyImportBase.Location
FROM
  FleetStudyImportBase
  LEFT JOIN vw_SixKeys ON FleetStudyImportBase.VIN = vw_SixKeys.VIN
WHERE
  (
    (
      (vw_SixKeys.VIN) Is Null
    )
  );
