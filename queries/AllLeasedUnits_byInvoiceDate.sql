SELECT
  vw_SixKeys.UnitInvoiceDate AS InvoiceDate,
  Format([AcceptedDate], "mmm") AS AcceptedMonth,
  Format([AcceptedDate], "yyyy") AS AcceptedYear,
  ClientGroups.ClientGroupName AS ClientGroup,
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
  vw_SixKeys.UnitDescShortVer AS UnitDesc,
  vw_SixKeys.BuildDate,
  vw_SixKeys.ActualDeliveryDate,
  vw_SixKeys.AcceptedDate,
  vw_SixKeys.InServiceDate,
  SchGrp.LeaseTermPrimary AS TermMonths,
  DateAdd(
    "m", [SchGrp].[LeaseTermPrimary],
    [SchGrp].[BLCD]
  ) AS LeaseToEnd,
  SchGrp.ExchangeITMonth,
  DateAdd(
    "m", [ExchangeITMonth], SchGrp.[BLCD]
  ) AS CanExchangeIT,
  vw_SixKeys.OffLeaseDate AS ActualOffLease,
  SchGrp.BLCD,
  vw_SixKeys.UnitCost,
  Units.UNITRENT AS SyndRent,
  Units.UNITSRENT AS Rent
FROM
  (
    (
      vw_SixKeys
      INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
    )
    INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
  )
  INNER JOIN ClientGroups ON vw_SixKeys.ClientGroupId = ClientGroups.ClientGroupId
WHERE
  (
    (
      (vw_SixKeys.UnitInvoiceDate) Between Forms!LORPTM!AnalSt
      And Forms!LORPTM!AnalEnd
    )
    And (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.UnitInvoiceDate DESC;
