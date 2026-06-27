SELECT
  ScheduleUnitDocs4.Lessee,
  ScheduleUnitDocs4.SchNo,
  ScheduleUnitDocs4.[Lease No],
  ScheduleUnitDocs4.Group,
  ScheduleUnitDocs4.[Unit No],
  ScheduleUnitDocs4.AssetId,
  ScheduleUnitDocs4.Status,
  ScheduleUnitDocs4.Unitvin,
  ScheduleUnitDocs4.Desc,
  ScheduleUnitDocs4.Location,
  ScheduleUnitDocs4.StreetAddress,
  ScheduleUnitDocs4.City,
  ScheduleUnitDocs4.State,
  ScheduleUnitDocs4.Zip,
  Cost_Refrig.RefrigMake,
  CCur(
    Nz([Cost_Refrig].[Refrig], 0)
  ) AS RefrigCost,
  Cost_Refrig.RefrigModel,
  ScheduleUnitDocs4.RefrigSerial,
  Units.UnitPlateNum AS PlateNum,
  Units.UnitInServInd AS InServiceInd,
  ScheduleUnitDocs4.InService AS InServiceDate,
  Units.UnitMSORecd AS MSORecv,
  Units.UnitMSOCopy AS MSOCopy,
  Units.UnitTASent AS TitleAppSent,
  [Last Title Entry].LogDate AS LastTitleLog,
  [Last Title Entry].Action AS TitleAction,
  ScheduleUnitDocs4.UnitCost,
  [TotalCost] - [UnitCost] AS ChildPartCost,
  ScheduleUnitDocs4.TotalCost AS Expr1
FROM
  (
    (
      ScheduleUnitDocs4
      LEFT JOIN Cost_Refrig ON ScheduleUnitDocs4.AssetId = Cost_Refrig.AssetId
    )
    LEFT JOIN Units ON ScheduleUnitDocs4.AssetId = Units.UnitID
  )
  LEFT JOIN [Last Title Entry] ON ScheduleUnitDocs4.AssetId = [Last Title Entry].AssetID
ORDER BY
  ScheduleUnitDocs4.Group,
  ScheduleUnitDocs4.[Unit No];
