SELECT
  Units.UnitUnitNum,
  Units.ClientID,
  Units.UnitID,
  Units.GroupREF,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.ObcUnitNumber,
  Units.unitstatus
FROM
  Units
WHERE
  (
    (
      (Units.UnitUnitNum) In (
        SELECT
          [UnitUnitNum]
        FROM
          [Units] As Tmp
        GROUP BY
          [UnitUnitNum]
        HAVING
          Count(*)> 1
      )
    )
    AND (
      (Units.ClientID) In (
        SELECT
          [ClientId]
        FROM
          [Units] As Tmp
        GROUP BY
          [ClientId]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  Units.UnitUnitNum;
