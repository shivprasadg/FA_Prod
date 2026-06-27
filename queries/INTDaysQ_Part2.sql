SELECT
  INTDaysQ_Part1.*,
  BillStEndq.Billstart,
  BillStEndq.BillEndDt,
  BillStEndq.SlvEnd,
  BillStEndq.PDEND,
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
    IsNull([billstart]),
    [ubstart],
    IIf(
      [ubstart] > [billstart], [ubstart],
      [billstart]
    )
  ) AS UBSlvS,
  IIf(
    [billenddt] < [slvend], [billenddt],
    [slvend]
  ) AS ubslvend
FROM
  (
    INTDaysQ_Part1
    INNER JOIN BillStEndq ON (
      INTDaysQ_Part1.SchNo = BillStEndq.SchNo
    )
    AND (
      INTDaysQ_Part1.MLNo = BillStEndq.MLNo
    )
  )
  INNER JOIN BillUnitq ON (
    INTDaysQ_Part1.MLNo = BillUnitq.MLNo
  )
  AND (
    INTDaysQ_Part1.SchNo = BillUnitq.SchNo
  );
