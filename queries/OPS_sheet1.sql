SELECT
  ScheduleUnitDocs2.Lessee,
  ScheduleUnitDocs2.SchNo,
  ScheduleUnitDocs2.[Lease No],
  ScheduleUnitDocs2.Group,
  ScheduleUnitDocs2.[Unit No],
  ScheduleUnitDocs2.AssetId,
  ScheduleUnitDocs2.Status,
  ScheduleUnitDocs2.Unitvin,
  ScheduleUnitDocs2.Desc,
  ScheduleUnitDocs2.Location,
  ScheduleUnitDocs2.StreetAddress,
  ScheduleUnitDocs2.City,
  ScheduleUnitDocs2.State,
  ScheduleUnitDocs2.Zip,
  Cost_Refrig.RefrigMake,
  CCur(
    Nz([Cost_Refrig].[Refrig], 0)
  ) AS RefrigCost,
  Cost_Refrig.RefrigModel,
  ScheduleUnitDocs2.RefrigSerial,
  Units.UnitPlateNum AS PlateNum,
  Units.UnitInServInd AS InServiceInd,
  ScheduleUnitDocs2.InService AS InServiceDate,
  Units.UnitMSORecd AS MSORecv,
  Units.UnitMSOCopy AS MSOCopy,
  Units.UnitTASent AS TitleAppSent,
  [Last Title Entry].LogDate AS LastTitleLog,
  [Last Title Entry].Action AS TitleAction,
  ScheduleUnitDocs2.UnitCost,
  [TotalCost] - [UnitCost] AS ChildPartCost,
  ScheduleUnitDocs2.TotalCost
FROM
  (
    (
      ScheduleUnitDocs2
      LEFT JOIN Cost_Refrig ON ScheduleUnitDocs2.AssetId = Cost_Refrig.AssetId
    )
    LEFT JOIN Units ON ScheduleUnitDocs2.AssetId = Units.UnitID
  )
  LEFT JOIN [Last Title Entry] ON ScheduleUnitDocs2.AssetId = [Last Title Entry].AssetID
ORDER BY
  ScheduleUnitDocs2.Group,
  ScheduleUnitDocs2.[Unit No];
