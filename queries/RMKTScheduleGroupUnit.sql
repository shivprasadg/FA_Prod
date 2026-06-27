SELECT
  Schedule.SchID,
  SchGrp.SGrpID AS GroupId,
  Units.UnitID AS AssetID,
  Format([rmkt].[BuyOutDate], "mm/dd/yy") AS BuyOutDate,
  Units.UnitUnitNum AS UnitNum,
  Units.UnitVIN AS VIN,
  Units.UnitStatus AS S,
  Banks.BankShortName AS Bank,
  Rmkt.IFLoanNumber AS [Loan#],
  Units.unitmake AS Make,
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
