SELECT
  Count(RMKTSalesPersons_Last30.ID) AS ID,
  Sum(
    RMKTSalesPersons_Last30.CallsIn
  ) AS CallsIn,
  Sum(
    RMKTSalesPersons_Last30.CallsOut
  ) AS CallsOut,
  Sum(
    RMKTSalesPersons_Last30.EmailsIn
  ) AS EmailsIn,
  Sum(
    RMKTSalesPersons_Last30.EmailsOut
  ) AS EmailsOut,
  Sum(
    RMKTSalesPersons_Last30.VisitsScheduled
  ) AS VisitsScheduled,
  Sum(
    RMKTSalesPersons_Last30.Proposals
  ) AS Proposals,
  Sum(
    RMKTSalesPersons_Last30.ProposalsActive
  ) AS ProposalsActive,
  Sum(
    RMKTSalesPersons_Last30.ProposalsSold
  ) AS ProposalsSold,
  Sum(RMKTSalesPersons_Last30.Sold30) AS Sold30,
  Format(
    Sum(
      [RMKTSalesPersons_Last30].[Sales30]
    ),
    "$#,##0"
  ) AS Sales30,
  Max(
    RMKTSalesPersons_Last30.LastCall
  ) AS LastCall,
  Max(
    RMKTSalesPersons_Last30.LastEmail
  ) AS LastEmail,
  Max(
    RMKTSalesPersons_Last30.LastVisit
  ) AS LastVisit,
  Max(
    RMKTSalesPersons_Last30.NextVisit
  ) AS NextVisit,
  Sum(
    RMKTSalesPersons_Last30.EmailsPending
  ) AS EmailsPending,
  Sum(
    RMKTSalesPersons_Last30.CallsPending
  ) AS CallsPending
FROM
  RMKTSalesPersons_Last30;
