SELECT
  BuyersProposal.BuyerID,
  Count(BuyersProposal.ProposalID) AS ProposalCount
FROM
  BuyersProposal
GROUP BY
  BuyersProposal.BuyerID;
