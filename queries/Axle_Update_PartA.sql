SELECT
  Units.UnitID,
  Units.UnitVIN,
  Units.Axle,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitYr,
  Left([UnitVIN], 11) AS VinPrefix
FROM
  Units
WHERE
  (
    (
      (Units.Axle)= "TBD"
    )
  );
