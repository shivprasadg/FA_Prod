SELECT
  Min(Deals.dealOutcomeDate) AS MinOfdealOutcomeDate,
  Deals.clientRef,
  First(Deals.DealLOEx) AS OrigBDEX,
  Deals.dealOutcome
FROM
  Deals
GROUP BY
  Deals.clientRef,
  Deals.dealOutcome
HAVING
  (
    (
      (Deals.dealOutcome)= "Awarded"
    )
  )
ORDER BY
  Deals.clientRef;
