SELECT
  PDiemBill_Part1.BillID,
  PDiemBill_Part1.clientID,
  PDiemBill_Part1.MLID,
  PDiemBill_Part1.SchID,
  PDiemBill_Part1.SGrpID,
  PDiemBill_Part1.UnitID,
  PDiemBill_Part1.MLNo,
  PDiemBill_Part1.SchNo,
  PDiemBill_Part1.SGrpNO,
  PDiemBill_Part1.BLCD,
  PDiemBill_Part1.SGrpNoSlv,
  PDiemBill_Part1.SGRPNoPD,
  PDiemBill_Part1.unitunitnum,
  PDiemBill_Part1.UnitVIN,
  PDiemBill_Part1.UnitStatus,
  PDiemBill_Part1.UnitAcceptDate,
  PDiemBill_Part1.Desc,
  PDiemBill_Part1.UnitTaxExempt,
  PDiemBill_Part1.clientCompanyName,
  PDiemBill_Part1.FAVenorNo,
  PDiemBill_Part1.SchBnkClNO,
  PDiemBill_Part1.vin4,
  PDiemBill_Part1.unitpdrent,
  PDiemBill_Part1.SGrpFSLPmt,
  PDiemBill_Part1.unitinserv,
  PDiemBill_Part1.schcrexp,
  PDiemBill_Part1.SchFSL,
  PDiemBill_Part1.unitrent,
  PDiemBill_Part1.unittax,
  PDiemBill_Part1.UNITOFFLEASEDT,
  PDiemBill_Part1.UnitPDRent,
  PDiemBill_Part1.SchCurrRent,
  PDiemBill_Part1.BillDueDt,
  PDiemBill_Part1.DueonRec,
  PDiemBill_Part1.Billdt,
  PDiemBill_Part1.BillSent,
  BillStEndq.Billstart,
  BillStEndq.BillEndDt,
  BillStEndq.SlvEnd,
  BillStEndq.PDEND,
  BillUnitq.UBStart,
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
  ) AS ubslvend
FROM
  (
    PDiemBill_Part1
    LEFT JOIN BillStEndq ON PDiemBill_Part1.BillID = BillStEndq.BillID
  )
  LEFT JOIN BillUnitq ON PDiemBill_Part1.BillID = BillUnitq.BillID
GROUP BY
  PDiemBill_Part1.BillID,
  PDiemBill_Part1.clientID,
  PDiemBill_Part1.MLID,
  PDiemBill_Part1.SchID,
  PDiemBill_Part1.SGrpID,
  PDiemBill_Part1.UnitID,
  PDiemBill_Part1.MLNo,
  PDiemBill_Part1.SchNo,
  PDiemBill_Part1.SGrpNO,
  PDiemBill_Part1.BLCD,
  PDiemBill_Part1.SGrpNoSlv,
  PDiemBill_Part1.SGRPNoPD,
  PDiemBill_Part1.unitunitnum,
  PDiemBill_Part1.UnitVIN,
  PDiemBill_Part1.UnitStatus,
  PDiemBill_Part1.UnitAcceptDate,
  PDiemBill_Part1.Desc,
  PDiemBill_Part1.UnitTaxExempt,
  PDiemBill_Part1.clientCompanyName,
  PDiemBill_Part1.FAVenorNo,
  PDiemBill_Part1.SchBnkClNO,
  PDiemBill_Part1.vin4,
  PDiemBill_Part1.unitpdrent,
  PDiemBill_Part1.SGrpFSLPmt,
  PDiemBill_Part1.unitinserv,
  PDiemBill_Part1.schcrexp,
  PDiemBill_Part1.SchFSL,
  PDiemBill_Part1.unitrent,
  PDiemBill_Part1.unittax,
  PDiemBill_Part1.UNITOFFLEASEDT,
  PDiemBill_Part1.UnitPDRent,
  PDiemBill_Part1.SchCurrRent,
  PDiemBill_Part1.BillDueDt,
  PDiemBill_Part1.DueonRec,
  PDiemBill_Part1.Billdt,
  PDiemBill_Part1.BillSent,
  BillStEndq.Billstart,
  BillStEndq.BillEndDt,
  BillStEndq.SlvEnd,
  BillStEndq.PDEND,
  BillUnitq.UBStart,
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
  ),
  IIf(
    [ubstart] > [billstart], [ubstart],
    [billstart]
  ),
  IIf(
    [billenddt] < [slvend], [billenddt],
    [slvend]
  );
