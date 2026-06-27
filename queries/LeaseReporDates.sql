SELECT
  vw_SixKeys.MLNo AS [Lease No],
  vw_SixKeys.CompanyName AS Lessee,
  vw_SixKeys.MLOrig,
  vw_SixKeys.MLDate AS MstrLeaseDate,
  vw_SixKeys.FAID,
  vw_SixKeys.PurchaseOrderDate,
  Deals.FAID AS DealFAID,
  Deals.dealOutcome,
  Deals.dealOutcomeDate,
  Deals.dealRALBack,
  Deals.dealRFPDate,
  Deals.dealLPAFcomp,
  Deals.dealLPAFDue,
  Deals.dealLPAFSent
FROM
  vw_SixKeys
  LEFT JOIN Deals ON vw_SixKeys.FAID = Deals.FAID
GROUP BY
  vw_SixKeys.MLNo,
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.MLDate,
  vw_SixKeys.FAID,
  vw_SixKeys.PurchaseOrderDate,
  Deals.FAID,
  Deals.dealOutcome,
  Deals.dealOutcomeDate,
  Deals.dealRALBack,
  Deals.dealRFPDate,
  Deals.dealLPAFcomp,
  Deals.dealLPAFDue,
  Deals.dealLPAFSent
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (Deals.dealOutcome)= "Awarded"
    )
  )
ORDER BY
  Deals.dealOutcomeDate DESC;
