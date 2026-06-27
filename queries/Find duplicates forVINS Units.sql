SELECT
  vw_SixKeys.SGrpID,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Units.UnitVIN,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitSubType,
  Units.ClientGroupId,
  Units.ClientID,
  Units.UnitStatus
FROM
  Units
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
WHERE
  (
    (
      (Units.UnitVIN) In (
        SELECT
          [UnitVIN]
        FROM
          [Units] As Tmp
        GROUP BY
          [UnitVIN]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName DESC,
  Units.UnitVIN;
