SELECT
  Units.UnitType,
  Units.UnitMake,
  Units.UnitModel,
  Count(Units.UnitID) AS Units
FROM
  Units
WHERE
  (
    (
      (Units.UnitStatus)<> "T"
    )
  )
GROUP BY
  Units.UnitType,
  Units.UnitMake,
  Units.UnitModel
ORDER BY
  Units.UnitType,
  Units.UnitMake,
  Units.UnitModel;
