SELECT
  MoveOrderImport.[Order#],
  MoveOrderImport.[Serial#],
  MoveOrderImport.[Driver#],
  MoveOrderImport.EntryDate,
  MoveOrderImport.LoadDate,
  MoveOrderImport.ETADate,
  MoveOrderImport.PreAssignedDate,
  MoveOrderImport.[Rev Terminal],
  MoveOrderImport.Status,
  MoveOrderImport.[Origin Dest],
  MoveOrderImport.[Shipper Name],
  MoveOrderImport.[Consignee Name],
  MoveOrderImport.Miles,
  MoveOrderImport.Commodity,
  MoveOrderImport.ELD,
  MoveOrderImport.TOWABLE,
  MoveOrderImport.InStorage
FROM
  [SELECT UnitID, UnitVIN, Right([UnitVin],
  8
) AS Veight
FROM
  Units
WHERE
  UnitStatus Like "E*" ] AS U
  RIGHT JOIN MoveOrderImport ON U.Veight = MoveOrderImport.[Serial#]
WHERE
  (
    (
      (U.UnitID) Is Null
    )
  );
