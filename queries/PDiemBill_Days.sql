SELECT
  BillStEndq.BillID,
  MstrLease.MLID,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpID,
  BillUnitq.UnitID,
  IIf(
    [billstart] > [schcrexp],
    0,
    IIf(
      [sgrpnoslv] = No,
      0,
      IIf(
        [ubstart] > [slvend],
        0,
        IIf(
          [ubstart] > [billenddt],
          0,
          DateDiff(
            "d",
            IIf(
              [ubstart] > [billstart], [ubstart],
              [billstart]
            ),
            IIf(
              [billenddt] < [slvend], [billenddt],
              [slvend]
            )+ 1
          )
        )
      )
    )
  ) AS SlvDays,
  IIf(
    [ubstart] > [billstart], [ubstart],
    [billstart]
  ) AS UBSlvS,
  IIf(
    [billenddt] < [slvend], [billenddt],
    [slvend]
  ) AS ubslvend,
  BillUnitq.ubstart,
  BillStEndq.slvend,
  BillStEndq.Billstart
FROM
  (
    MstrLease
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN (
    (
      SchGrp
      INNER JOIN BillStEndq ON SchGrp.SGrpID = BillStEndq.SGrpID
    )
    INNER JOIN BillUnitq ON BillStEndq.SGrpID = BillUnitq.SGrpID
  ) ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (BillStEndq.BillID)= [Forms]![Billing]![BillingSubForm].[Form]![BillID]
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchID;
