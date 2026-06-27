SELECT
  Units.UnitID,
  Units.UnitVIN,
  Units.UnitInServ,
  Units.UnitAcceptDate,
  Units.UnitDocAccpt,
  [UnitAcceptDate] - [UnitInServ] AS DaysDiff,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType
FROM
  Units
WHERE
  (
    (
      (Units.UnitAcceptDate)> [UnitInServ]
    )
  )
ORDER BY
  Units.UnitInServ DESC;
