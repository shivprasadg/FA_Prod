SELECT
  Units.UnitGarageLocationRef AS GarageLocationID,
  Units.UnitID AS AssetID,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  Units.UnitUnitNum AS [Unit#],
  [UnitYr] & " " & [UnitMake] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& " " & [UnitType] AS [Desc],
  Units.UnitYr AS UnitYear,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  Units.UnitSubType,
  Units.Axle,
  Units.UnitStatus,
  IIf(
    [unitstatus] = "A"
    Or [unitstatus] = "B"
    Or [unitstatus] = "R"
    Or [unitstatus] = "T"
    Or [unitstatus] = "TR",
    1, 0
  ) AS InFleet,
  Units.ClientGroupId,
  Units.ClientID,
  vw_SixKeys.MLOrig
FROM
  Units
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
ORDER BY
  vw_SixKeys.MLOrig,
  Units.UnitID;
