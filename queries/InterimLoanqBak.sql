SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.LoanBank,
  SchGrp.LoanDate,
  SchGrp.LoanAmount,
  SchGrp.LoanType,
  SchGrp.SGrpExpDt,
  SchGrp.SgrpPayoffDt,
  Units.UnitVIN,
  Units.unitstatus,
  [unitfinalcost] + IIf(
    IsNull([sumofchildfinalcost]),
    0,
    [sumofchildfinalcost]
  ) AS Cost,
  Units.UnitIFDate,
  Units.UnitIFBank,
  Units.UnitIFAmt,
  Units.UnitPayOff,
  [FORMS]![LOANPARMF]![ASOFDT] AS [ASOF DT],
  SchGrp.SgrpPayoffDt
FROM
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
  INNER JOIN (
    Units
    LEFT JOIN ChildPartSummaryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (SchGrp.LoanBank) Is Not Null
    )
    And (
      (SchGrp.LoanDate) Is Null
      Or (SchGrp.LoanDate)< forms!loanparmf!asofdt
    )
    And (
      (SchGrp.LoanAmount)> 0.001
    )
    And (
      (SchGrp.LoanType)= forms!loanparmf!loantype
    )
    And (
      (SchGrp.SgrpPayoffDt) Is Null
      Or (SchGrp.SgrpPayoffDt)> forms!loanparmf!asofdt
    )
    And (
      (Clients.ClientGroupID)<> 18
    )
  )
  Or (
    (
      (SchGrp.LoanAmount)> 0.001
    )
    And (
      (SchGrp.LoanType)= forms!loanparmf!loantype
    )
    And (
      (Units.UnitIFDate) Is Not Null
      Or (Units.UnitIFDate)< forms!LOANPARMF!ASOFDT
    )
    And (
      (Clients.ClientGroupID)<> 18
    )
  )
ORDER BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpNO;
