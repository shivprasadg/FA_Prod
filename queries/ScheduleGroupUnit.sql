SELECT
  Schedule.SchID,
  SchGrp.SGrpID AS GroupId,
  Units.UnitID AS AssetID,
  Units.UnitUnitNum AS UnitNum,
  Units.UnitVIN AS VIN,
  Units.UnitStatus AS Status,
  Banks.BankShortName AS Bank,
  Rmkt.IFLoanNumber AS [Loan#],
  Units.unitmake AS Make,
  Units.Unityr AS MYear
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
ORDER BY
  Units.UnitID;
