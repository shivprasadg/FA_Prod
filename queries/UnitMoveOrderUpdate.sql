SELECT
  U.UnitID,
  MoveOrderImport.[Order#],
  MoveOrderImport.EntryDate,
  MoveOrderImport.LoadDate,
  MoveOrderImport.ETADate,
  MoveOrderImport.Status,
  MoveOrderImport.Miles
FROM
  (
    SELECT
      UnitID,
      UnitVIN,
      Right([UnitVin], 8) AS Veight
    FROM
      Units
    WHERE
      UnitStatus = "EM"
  ) AS U
  INNER JOIN MoveOrderImport ON U.Veight = MoveOrderImport.[Serial#];
