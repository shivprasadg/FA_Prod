SELECT
  DISTINCT BuyersProposal.SalesPersonID,
  BuyersProposalUnits.ProposalID,
  BuyersProposalUnits.UnitID
FROM
  BuyersProposalUnits
  INNER JOIN BuyersProposal ON BuyersProposalUnits.ProposalID = BuyersProposal.ProposalID;
