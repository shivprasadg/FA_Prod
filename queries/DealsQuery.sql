SELECT
  "Edit" AS edit,
  Deals.dealID,
  lookup_transTypes.TransType,
  Deals.clientRef,
  Deals.FAID,
  [clientcompanyName] & IIf(
    [clientDBA] <> "", " d/b/a " & [clientDBA],
    ""
  ) AS companyAndDBA,
  Deals.Increment,
  Deals.dealShortDesc,
  Deals.dealRFPDate,
  Deals.dealLPAFDue,
  Deals.dealDelDate,
  Deals.dealDelThru,
  Deals.CRfinancialsIH,
  Deals.CRFinancialLink,
  Deals.CRGuarantor,
  Deals.compBank,
  Deals.dealLPAFSent,
  Deals.dealOutcomeDate,
  Deals.dealOutcome,
  Deals.dealOutcomeReason,
  Deals.dealUnitQty,
  Deals.dealUnitPrice,
  Deals.dealGrandTotal,
  Deals.dealSchedNum,
  Deals.transTypeRef,
  Deals.dealEnteredDate,
  Deals.dealEstDelDate,
  Deals.dealMPY,
  Deals.dealTerm,
  Deals.dealGrandTotal,
  Deals.dealLPAFcomp,
  Deals.LostTo,
  Deals.LostRate,
  Deals.LostTerm,
  Deals.LostResid,
  Deals.[LostImplicit Rt],
  Deals.LostTo,
  Deals.DealLOEx
FROM
  Clients
  INNER JOIN (
    Deals
    INNER JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
  ) ON Clients.clientID = Deals.clientRef;
