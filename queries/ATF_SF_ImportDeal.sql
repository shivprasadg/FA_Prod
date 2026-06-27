SELECT
  vw_ATF_Avail_Signed_FAID.Deals_FAID AS FAID,
  Sum(
    vw_ATF_Avail_Signed_FAID.Deals_dealGrandTotal
  ) AS SumOfDeals_dealGrandTotal,
  vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc AS dealShortDesc,
  vw_ATF_Avail_Signed_FAID.Deals_RFPDate AS dealRFPDate,
  vw_ATF_Avail_Signed_FAID.Deals_dealLPAFDue AS dealLPAFDue,
  vw_ATF_Avail_Signed_FAID.Deals_dealLPAFcomp AS dealLPAFcomp,
  Sum(
    vw_ATF_Avail_Signed_FAID.Deals_dealUnitQty
  ) AS dealUnitQty,
  [Forms]![ATF_SFImporter]![LesseeID] AS clientRef,
  'Awarded' AS dealOutcome,
  CDate([LPAFDate]) AS dealOutcomeDate,
  vw_ATF_Avail_Signed_FAID.Deals_dealDateEntered AS dealEnteredDate,
  Max(
    vw_ATF_Avail_Signed_FAID.Deals_dealEstDelDate
  ) AS dealEstDelDate,
  Min(
    vw_ATF_Avail_Signed_FAID.Deals_dealMPY
  ) AS dealMPY,
  Max(
    vw_ATF_Avail_Signed_FAID.Deals_dealTerm
  ) AS dealTerm,
  Avg(
    vw_ATF_Avail_Signed_FAID.Deals_dealUnitPrice
  ) AS dealUnitPrice,
  Sum(
    vw_ATF_Avail_Signed_FAID.Deals_dealGrandTotal
  ) AS dealGrandTotal,
  Salesperson.SPInitials AS DealLOEx,
  vw_ATF_Avail_Signed_FAID.Type
FROM
  vw_ATF_Avail_Signed_FAID
  INNER JOIN Salesperson ON vw_ATF_Avail_Signed_FAID.SalespersonID = Salesperson.SalespersonID
GROUP BY
  vw_ATF_Avail_Signed_FAID.Deals_FAID,
  vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc,
  vw_ATF_Avail_Signed_FAID.Deals_RFPDate,
  vw_ATF_Avail_Signed_FAID.Deals_dealLPAFDue,
  vw_ATF_Avail_Signed_FAID.Deals_dealLPAFcomp,
  CDate([LPAFDate]),
  vw_ATF_Avail_Signed_FAID.Deals_dealDateEntered,
  Salesperson.SPInitials,
  vw_ATF_Avail_Signed_FAID.Type
ORDER BY
  vw_ATF_Avail_Signed_FAID.Deals_FAID;
