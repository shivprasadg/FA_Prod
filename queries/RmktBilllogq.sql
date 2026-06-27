SELECT
  Clients.clientCompanyName,
  Clients.ClientShNm,
  RmktBills.[Invoice No],
  RmktBills.RevNo,
  RmktBills.Billdt,
  RmktBills.RBillDueDt,
  RmktBills.BillSent,
  Rmkt.InvoiceRequestDate,
  Count(RmktBills.ID) AS CountOfID,
  Rmkt.InvoicePaidDate,
  Units.unitstatus,
  IIf(
    [unitstatus] = "s",
    3,
    IIf(
      [unitstatus] = "ks",
      4,
      IIf(
        IsNull([billsent]),
        1,
        2
      )
    )
  ) AS BillSt,
  Sum(Rmkt.SoldPrice) AS SumOfSoldPrice
FROM
  (
    Clients
    INNER JOIN RmktBills ON Clients.clientID = RmktBills.BuyerRef
  )
  LEFT JOIN (
    Units
    RIGHT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  ) ON (
    RmktBills.ReqDt = Rmkt.InvoiceRequestDate
  )
  AND (
    RmktBills.BuyerRef = Rmkt.BuyerRef
  )
WHERE
  (
    (
      (Clients.clientType)= "remarketing"
    )
  )
GROUP BY
  Clients.clientCompanyName,
  Clients.ClientShNm,
  RmktBills.[Invoice No],
  RmktBills.RevNo,
  RmktBills.Billdt,
  RmktBills.RBillDueDt,
  RmktBills.BillSent,
  Rmkt.InvoiceRequestDate,
  Rmkt.InvoicePaidDate,
  Units.unitstatus,
  IIf(
    [unitstatus] = "s",
    3,
    IIf(
      [unitstatus] = "ks",
      4,
      IIf(
        IsNull([billsent]),
        1,
        2
      )
    )
  ),
  RmktBills.ID
ORDER BY
  Clients.ClientShNm;
