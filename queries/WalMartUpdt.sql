UPDATE
  Units
  INNER JOIN Walmart ON Units.UnitID = Walmart.UnitID
SET
  Units.unitunitnum = [unit #],
  Units.unitvin = [vin]
WHERE
  (
    (
      (Units.UnitID) Is Not Null
    )
  );
