SELECT
  Units.*,
  Locations.LState
FROM
  (
    Units
    INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID
  )
  LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID;
