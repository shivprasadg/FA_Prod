SELECT
  DISTINCT Bills.BillID,
  MstrLease.MLNo,
  Schedule.SchNo,
  Schedule.SchID,
  SchGrp.SGrpID,
  IIf(
    IsNull([BLCD]),
    [Forms]![Billing]![BillingSubForm].[Form]![billto],
    IIf(
      [Forms]![Billing]![BillingSubForm].[Form]![billto] > [BLCD],
      [BLCD], [Forms]![Billing]![BillingSubForm].[Form]![billto]
    )
  ) AS BillEndDt,
  Bills.Billfrom AS Billstart,
  IIf(
    [billto] > IIf(
      IsNull([BLCD]),
      [schcrexp],
      IIf(
        [schcrexp] > [BLCD], [BLCD], [schcrexp]
      )
    ),
    IIf(
      IsNull([BLCD]),
      [schcrexp],
      IIf(
        [schcrexp] > [BLCD], [BLCD], [schcrexp]
      )
    ),
    [billto]
  ) AS SlvEnd,
  IIf(
    IsNull([BLCD]),
    [billto],
    IIf(
      [billto] > [BLCD], [BLCD], [billto]
    )
  ) AS PDEND
FROM
  Bills
  INNER JOIN (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  ) ON Bills.BillSchRef = Schedule.SchID
WHERE
  (
    (
      (Bills.BillID)= [Forms]![Billing]![BillingSubForm].[Form]![BillID]
    )
  );
