SELECT
  Units.ObcUnitNumber,
  Units.ClientGroupId,
  Units.UnitID,
  Units.UnitYr,
  Units.ClientID,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType
FROM
  Units
WHERE
  (
    (
      (Units.ObcUnitNumber) In (
        SELECT
          [ObcUnitNumber]
        FROM
          [Units] As Tmp
        GROUP BY
          [ObcUnitNumber],
          [ClientGroupId]
        HAVING
          Count(*)> 1
          And [ClientGroupId] = [Units].[ClientGroupId]
      )
    )
    AND (
      (Units.ClientGroupId)= 697
    )
  )
ORDER BY
  Units.ObcUnitNumber,
  Units.ClientGroupId,
  Units.UnitYr DESC,
  Units.ClientID;
