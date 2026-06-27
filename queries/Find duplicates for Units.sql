SELECT
  Units.ClientGroupId,
  Units.ObcUnitNumber,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitVIN,
  Units.ClientID
FROM
  Units
WHERE
  (
    (
      (Units.ClientGroupId) In (
        SELECT
          [ClientGroupId]
        FROM
          [Units] As Tmp
        GROUP BY
          [ClientGroupId],
          [ObcUnitNumber]
        HAVING
          Count(*)> 1
          And [ObcUnitNumber] = [Units].[ObcUnitNumber]
      )
    )
  )
ORDER BY
  Units.ClientGroupId,
  Units.ObcUnitNumber;
