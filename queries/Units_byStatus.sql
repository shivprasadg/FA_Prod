SELECT
  Units.UnitStatus,
  Count(Units.UnitID) AS UnitCount,
  UnitStatus.DEF
FROM
  Units
  INNER JOIN UnitStatus ON Units.UnitStatus = UnitStatus.Unitstatus
GROUP BY
  Units.UnitStatus,
  UnitStatus.DEF
ORDER BY
  Units.UnitStatus;
