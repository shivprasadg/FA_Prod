UPDATE
  [UP lic]
  INNER JOIN Units ON [UP lic].[Unit ID] = Units.UnitID
SET
  Units.UnitPlateNum = [license];
