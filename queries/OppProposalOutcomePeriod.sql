SELECT
  Clients.LOinitial,
  Clients.ClientShNm AS Client,
  Deals.FAID,
  DatePart('q', [dealOutcomeDate]) AS DealQTR,
  DatePart('yyyy', [dealOutcomeDate]) AS DealYear,
  Deals.dealOutcome,
  Deals.dealLPAFSent AS LPAFSent,
  Deals.dealOutcomeDate,
  Round([deals].[dealGrandTotal]) AS dealGrandTotal,
  Deals.dealShortDesc,
  Deals.dealoutcome AS RptGrp,
  Deals.DealLOEx,
  lookup_transTypes.TransType,
  IIf(
    [DealLOEx] = "hs"
    Or [DealLOEx] = "sm",
    "HS", [DealLOEx]
  ) AS RptLOEx,
  Deals.dealUnitQty,
  Deals.dealSchedNum,
  Clients.clienttype,
  "x" AS Grpon,
  IIf(
    [dealoutcome] = "awarded",
    1,
    IIf([dealoutcome] = "lost", 2, 3)
  ) AS Outcomsrt,
  DatePart(
    "yyyy",
    Date()
  ) AS yr,
  Clients.ClientGroupID
FROM
  (
    Clients
    INNER JOIN Salesperson ON Clients.LOinitial = Salesperson.SPInitials
  )
  INNER JOIN (
    Deals
    INNER JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
  ) ON Clients.clientID = Deals.clientRef
WHERE
  (
    (
      (Clients.ClientGroupID)<> 18
    )
  )
ORDER BY
  Clients.LOinitial,
  Clients.ClientShNm,
  Deals.dealOutcome;
