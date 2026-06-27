SELECT
  Clients.clientID,
  Deals.FAID,
  Deals.DealLOEx,
  Clients.clientBecameDate,
  Max(Deals.dealOutcomeDate) AS dealOutcomeDate,
  DateDiff(
    "m", [clientBecameDate], [dealOutcomeDate]
  ) AS ClientAgeAtDeal,
  IIf(
    DateDiff(
      "m", [clientBecameDate], [dealOutcomeDate]
    )> 12,
    CDbl(0.08),
    CDbl(0.12)
  ) AS DealCommission
FROM
  Clients
  LEFT JOIN Deals ON Clients.clientID = Deals.clientRef
WHERE
  (
    (
      (Deals.dealOutcome)= "Awarded"
    )
  )
GROUP BY
  Clients.clientID,
  Deals.FAID,
  Deals.DealLOEx,
  Clients.clientBecameDate,
  IIf(
    DateDiff(
      "m", [clientBecameDate], [dealOutcomeDate]
    )> 12,
    CDbl(0.08),
    CDbl(0.12)
  )
HAVING
  (
    (
      (
        Max(Deals.dealOutcomeDate)
      ) Is Not Null
    )
  )
ORDER BY
  Clients.clientID,
  Max(Deals.dealOutcomeDate);
