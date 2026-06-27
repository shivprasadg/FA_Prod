SELECT
  Units.UnitUnitNum,
  Units.ClientGroupId,
  Units.UnitID,
  Units.GroupREF,
  Units.UnitStatus,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitVIN,
  Units.clientID,
  Units.UnitSubType,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.MLOrig,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup
FROM
  Units
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitId
WHERE
  (
    (
      (Units.UnitUnitNum) In (
        SELECT
          [UnitUnitNum]
        FROM
          [Units] As Tmp
        GROUP BY
          [UnitUnitNum],
          [ClientGroupId]
        HAVING
          Count(*)> 1
          And [ClientGroupId] = [Units].[ClientGroupId]
      )
    )
    AND (
      (Units.ClientGroupId)<> 91
      And (Units.ClientGroupId)<> 53
      And (Units.ClientGroupId)<> 50
    )
  )
ORDER BY
  Units.UnitUnitNum,
  Units.ClientGroupId;
