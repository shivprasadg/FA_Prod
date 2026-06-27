SELECT
  Deals.FAID,
  Clients.clientcompanyname AS Client,
  DatePart("yyyy", [DealLPAFSent]) AS SentYear,
  DatePart("yyyy", [dealoutcomedate]) AS DealYear,
  Sum(Deals.dealUnitQty) AS DealUnitQty,
  Sum(Deals.dealGrandTotal) AS DealGrandTotal,
  Max(Deals.dealOutcomeDate) AS OutcomeDate,
  Deals.transTypeRef,
  Deals.DealLOEx AS DealBEDEX,
  Clients.ClientGroupId
FROM
  Clients
  INNER JOIN Deals ON Clients.clientID = Deals.clientRef
WHERE
  (
    (
      (Deals.dealOutcome)= "Awarded"
    )
    AND (
      (Clients.ClientGroupId)<> 18
    )
  )
GROUP BY
  Deals.FAID,
  Clients.clientcompanyname,
  DatePart("yyyy", [DealLPAFSent]),
  DatePart("yyyy", [dealoutcomedate]),
  Deals.transTypeRef,
  Deals.DealLOEx,
  Clients.ClientGroupId
ORDER BY
  Clients.clientcompanyname,
  Deals.FAID;
