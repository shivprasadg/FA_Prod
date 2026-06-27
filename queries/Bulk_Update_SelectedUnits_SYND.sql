SELECT
  Units.*,
  Bulk_Update_WorkTable.UnitDesc,
  Bulk_Update_WorkTable.ChildPartCost,
  [UnitFinalCost] + Nz([ChildPartCost], 0) AS UnitOEM,
  Locations.LState
FROM
  (
    Units
    INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID
  )
  LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID;
