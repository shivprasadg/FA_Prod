SELECT
  Deals.transTypeRef,
  Deals.pipelineOrder,
  Deals.FAID,
  Deals.dealID,
  Deals.clientRef,
  [clientcompanyName] & IIf(
    [clientDBA] <> "", " d/b/a " & [clientDBA],
    ""
  ) AS CompanyAndDBA,
  Clients.clientCompanyName,
  Deals.dealEnteredDate,
  Deals.dealEstDelDate,
  Deals.dealOutcome,
  Deals.dealOutcomeDate,
  "View/Edit" AS Edit,
  Deals.dealShortDesc,
  Deals.dealRFPDate,
  Deals.dealLPAFDue,
  Deals.dealLPAFSent,
  Deals.ProposalTimeline,
  Deals.dealOutcomeReason,
  Deals.dealSchedNum,
  Deals.ExcludeFromKPI,
  Deals.DealLOEx
FROM
  Clients
  INNER JOIN Deals ON Clients.clientID = Deals.clientRef
WHERE
  (
    (
      (Deals.transTypeRef)<> 1
    )
  )
ORDER BY
  Deals.FAID DESC,
  Deals.dealID DESC;
