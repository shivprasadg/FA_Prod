SELECT
  Bills.BillID,
  Clients.clientID,
  MstrLease.MLID,
  Schedule.SchID,
  SchGrp.SGrpID,
  Units.UnitID,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.BLCD,
  SchGrp.SGrpNoSlv,
  SchGrp.SGRPNoPD,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.UnitStatus,
  Units.UnitAcceptDate,
  [UnitYr] & " " & [UnitMake] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& IIf(
    [UnitType] = "Tractor", "", " " & [UnitType]
  ) AS [Desc],
  Units.UnitTaxExempt,
  Clients.clientCompanyName,
  Clients.FAVenorNo,
  Schedule.SchBnkClNO,
  Right([unitvin], 6) AS vin4,
  Units.unitpdrent,
  SchGrp.SGrpFSLPmt,
  Units.unitinserv,
  Schedule.schcrexp,
  Schedule.SchFSL,
  Units.unitrent,
  Units.unittax,
  Units.UNITOFFLEASEDT,
  Units.UnitPDRent,
  Schedule.SchCurrRent,
  Bills.BillDueDt,
  Bills.DueonRec,
  Bills.Billdt,
  Bills.BillSent
FROM
  Bills
  INNER JOIN (
    (
      (
        (
          Clients
          INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
        )
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  ) ON Bills.BillSchRef = Schedule.SchID
GROUP BY
  Bills.BillID,
  Clients.clientID,
  MstrLease.MLID,
  Schedule.SchID,
  SchGrp.SGrpID,
  Units.UnitID,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.BLCD,
  SchGrp.SGrpNoSlv,
  SchGrp.SGRPNoPD,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.UnitStatus,
  Units.UnitAcceptDate,
  [UnitYr] & " " & [UnitMake] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& IIf(
    [UnitType] = "Tractor", "", " " & [UnitType]
  ),
  Units.UnitTaxExempt,
  Clients.clientCompanyName,
  Clients.FAVenorNo,
  Schedule.SchBnkClNO,
  Right([unitvin], 6),
  Units.unitpdrent,
  SchGrp.SGrpFSLPmt,
  Units.unitinserv,
  Schedule.schcrexp,
  Schedule.SchFSL,
  Units.unitrent,
  Units.unittax,
  Units.UNITOFFLEASEDT,
  Units.UnitPDRent,
  Schedule.SchCurrRent,
  Bills.BillDueDt,
  Bills.DueonRec,
  Bills.Billdt,
  Bills.BillSent
HAVING
  (
    (
      (Bills.BillID)= Forms!Billing!BillingSubForm.Form!BillID
    )
    And (
      (SchGrp.SGrpNoSlv) Is Not Null
    )
    And (
      (SchGrp.SGRPNoPD) Is Not Null
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  Units.unitunitnum,
  Right([unitvin], 6);
