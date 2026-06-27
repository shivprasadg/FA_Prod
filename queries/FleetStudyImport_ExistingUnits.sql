SELECT
  FleetStudyImportBase.UnitNumber,
  Units.UnitUnitNum AS AtlaasUnitNum,
  Units.ObcUnitNumber,
  FleetStudyImportBase.UnitYear,
  Units.UnitYr AS AtlaasYear,
  FleetStudyImportBase.VIN,
  FleetStudyImportBase.Type,
  FleetStudyImportBase.UnitMake,
  Units.UnitMake AS AtlaasMake,
  vw_SixKeys.MLNo,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  Units.GroupREF
FROM
  (
    FleetStudyImportBase
    INNER JOIN Units ON FleetStudyImportBase.VIN = Units.UnitVIN
  )
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
WHERE
  (
    (
      (Units.UnitVIN) Is Not Null
    )
  );
