SELECT
  Deals.FAID,
  Clients.clientcompanyname AS Client,
  DatePart("q", [DealLPAFSent]) AS LPAFQtr,
  DatePart("yyyy", [DealLPAFSent]) AS LPAFYear,
  DatePart("q", [dealoutcomedate]) AS DealQtr,
  DatePart("yyyy", [dealoutcomedate]) AS DealYear,
  Sum(Deals.dealUnitQty) AS DealUnitCount,
  Sum(Deals.dealGrandTotal) AS DealGrandTotal,
  Clients.clientID
FROM
  Clients
  INNER JOIN Deals ON Clients.clientID = Deals.clientRef
WHERE
  (
    (
      (Deals.dealOutcomeDate) Is Not Null
    )
    AND (
      (Deals.dealOutcome)= "Awarded"
    )
  )
GROUP BY
  Deals.FAID,
  Clients.clientcompanyname,
  DatePart("q", [DealLPAFSent]),
  DatePart("yyyy", [DealLPAFSent]),
  DatePart("q", [dealoutcomedate]),
  DatePart("yyyy", [dealoutcomedate]),
  Clients.clientID
ORDER BY
  Clients.clientcompanyname,
  Deals.FAID;
