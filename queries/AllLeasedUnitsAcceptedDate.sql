SELECT
  vw_SixKeys.UnitInvoiceDate AS InvoiceDate,
  Format([AcceptedDate], "mmm") AS AcceptedMonth,
  Format([AcceptedDate], "yyyy") AS AcceptedYear,
  vw_SixKeys.CompanyName,
  vw_SixKeys.FAID,
  vw_SixKeys.MLNo,
  vw_SixKeys.MLDate,
  vw_SixKeys.BankShortName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.PurchaseOrder,
  vw_SixKeys.PurchaseOrderDate,
  vw_SixKeys.UnitID,
  vw_SixKeys.Status,
  vw_SixKeys.Unitnum,
  vw_SixKeys.MYear,
  vw_SixKeys.Model,
  vw_SixKeys.Type,
  vw_SixKeys.BuildDate,
  vw_SixKeys.ActualDeliveryDate,
  vw_SixKeys.AcceptedDate,
  vw_SixKeys.InServiceDate,
  SchGrp.LeaseTermPrimary AS TermMonths,
  DateAdd("m", [LeaseTermPrimary], [BLCD]) AS LeaseToEnd,
  SchGrp.ExchangeITMonth,
  DateAdd("m", [ExchangeITMonth], [BLCD]) AS CanExchangeIT,
  vw_SixKeys.OffLeaseDate AS ActualOffLease,
  SchGrp.BLCD,
  vw_SixKeys.UnitCost,
  Units.UNITRENT AS Rent
FROM
  (
    vw_SixKeys
    INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
  )
  INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
WHERE
  (
    (
      (vw_SixKeys.AcceptedDate) Between [Forms]![LORPTM]![AnalSt]
      And [Forms]![LORPTM]![AnalEnd]
    )
  );
