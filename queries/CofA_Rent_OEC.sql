SELECT
  UnitOEC2.AssetId,
  UnitOEC2.TotalCost,
  UnitOEC2.PerDiemRent,
  UnitOEC2.SYNDRent,
  UnitOEC2.LORent
FROM
  Bulk_Update_WorkTable
  INNER JOIN UnitOEC2 ON Bulk_Update_WorkTable.UnitID = UnitOEC2.AssetId;
