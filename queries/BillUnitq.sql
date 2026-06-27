SELECT
  MstrLease.Client,
  Bills.BillID,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpID,
  Units.UnitID,
  IIf(
    IsNull([unitpdstartdate])
    And IsNull([unitinserv]),
    [unitacceptdate],
    Nz([unitpdstartdate], [unitinserv])
  ) AS UBStart,
  SchGrp.BLCD AS unitmax,
  Units.UnitInServInd,
  Units.UnitStatus,
  Units.UnitDocAccpt,
  IIf(
    [sgrpnoslv] = No,
    IIf(
      [UBSTART] > [BILLTO],
      [BILLTO] + 1,
      IIf(
        [UBSTART] > [BILLFROM], [UBSTART],
        [BILLFROM]
      )
    ),
    IIf(
      [ubstart] >= [schcrexp] + 1
      And [ubstart] >= [billfrom],
      [ubstart],
      IIf(
        [schcrexp] + 1 >= [billfrom], [schcrexp] + 1,
        [billfrom]
      )
    )
  ) AS UnitPDSTOld,
  IIf(
    [UBStart] >= [PSent]
    And [UBStart] <= [PStart]
    And [Client] = 4971,
    [UBStart],
    IIf(
      [sgrpnoslv] = No,
      IIf(
        [UBSTART] > [BILLTO],
        [BILLTO] + 1,
        IIf(
          [UBSTART] > [BILLFROM], [UBSTART],
          [BILLFROM]
        )
      ),
      IIf(
        [ubstart] >= [schcrexp] + 1
        And [ubstart] >= [billfrom],
        [ubstart],
        IIf(
          [schcrexp] + 1 >= [billfrom], [schcrexp] + 1,
          [billfrom]
        )
      )
    )
  ) AS UnitPDST,
  Units.UnitInServ,
  Bills.Billfrom,
  Schedule.SchID,
  CDate(
    Nz([BillDx], #1/1/1900#)
  ) AS PBillDate,
  CDate(
    Nz([Sent], #1/1/1910#)
  ) AS PSent,
  CDate(
    Nz([Start], #1/1/4500#)
  ) AS PStart,
  CDate(
    Nz([End], #12/31/4500#)
  ) AS PEnd
FROM
  (
    (
      MstrLease
      INNER JOIN (
        (
          Schedule
          INNER JOIN Bills ON Schedule.SchID = Bills.BillSchRef
        )
        INNER JOIN SchGrp ON Bills.BillSchRef = SchGrp.SGrpSchID
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  LEFT JOIN PerDiemBill_MostRecentPD2 ON Schedule.SchID = PerDiemBill_MostRecentPD2.SchID
WHERE
  (
    (
      (Bills.BillID)= [Forms]![Billing]![BillingSubForm].[Form]![BillID]
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpID,
  Units.UnitID;
