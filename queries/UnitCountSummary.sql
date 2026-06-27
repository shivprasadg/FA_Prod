SELECT
  Units.GroupREF,
  Count(Units.UnitID) AS CountOfUnitID
FROM
  Units
GROUP BY
  Units.GroupREF;
