SELECT
  Units.ClientGroupID,
  Units.ClientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLNo,
  Units.UnitID,
  Units.UnitVIN,
  Units.UnitUnitNum,
  Units.OBCUnitNumber,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  Units.UnitSubType,
  Units.Unitstatus
FROM
  (
    Units
    INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
  )
  LEFT JOIN Reporting_Units_CurrMonth ON Units.UnitID = Reporting_Units_CurrMonth.UnitID
WHERE
  (
    (
      (Units.ClientGroupID)= [Ask ClientGroupID]
    )
    AND (
      (Units.UnitType)= "daycab"
      Or (Units.UnitType)= "sleeper"
      Or (Units.UnitType) Like "*Straight*"
    )
    AND (
      (Units.Unitstatus)= "A"
      Or (Units.Unitstatus)= "T"
      Or (Units.Unitstatus)= "TR"
    )
    AND (
      (
        Reporting_Units_CurrMonth.UnitID
      ) Is Null
    )
  )
ORDER BY
  Units.ClientID,
  Units.UnitID,
  Units.OBCUnitNumber;
