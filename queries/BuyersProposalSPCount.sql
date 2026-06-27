SELECT
  BuyersProposal.SalesPersonID,
  Count(BuyersProposal.ProposalID) AS ProposalCount
FROM
  BuyersProposal
GROUP BY
  BuyersProposal.SalesPersonID;
