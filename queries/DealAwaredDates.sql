SELECT
  Deals.FAID,
  Deals.dealLPAFSent,
  Deals.dealOutcome,
  DatePart("yyyy", [dealOutcomeDate]) AS DealYear
FROM
  Deals
GROUP BY
  Deals.FAID,
  Deals.dealLPAFSent,
  Deals.dealOutcome,
  DatePart("yyyy", [dealOutcomeDate])
HAVING
  (
    (
      (Deals.dealLPAFSent) Is Not Null
    )
    AND (
      (Deals.dealOutcome)= "awarded"
    )
  )
ORDER BY
  Deals.dealLPAFSent DESC;
