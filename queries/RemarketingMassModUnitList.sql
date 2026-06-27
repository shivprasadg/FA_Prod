SELECT
  Schedule.SchID,
  SchGrp.SGrpID AS GroupId,
  Units.UnitID AS AssetID,
  Units.UnitUnitNum AS [Unit#],
  Units.UnitVIN AS VIN,
  Units.UnitStatus AS S,
  Units.UnitMake AS Make,
  Units.UnitSubType AS CabType,
  Units.UnitYr,
  SchGrp.SGrpNO AS UnitGroup
FROM
  (
    (
      Schedule
      INNER JOIN (
        SchGrp
        INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
      ) ON Schedule.SchID = SchGrp.SGrpSchID
    )
    LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  )
  LEFT JOIN Banks ON Rmkt.IFBankID = Banks.BankId
WHERE
  (
    (
      (Units.UnitStatus)<> "T"
      And (Units.UnitStatus)<> "O"
    )
  )
ORDER BY
  Units.UnitID;
