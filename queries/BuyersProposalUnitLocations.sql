SELECT
  Buyer2Proposal_Query.LocationID,
  Buyer2Proposal_Query.SourceCompany,
  Buyer2Proposal_Query.Location,
  Count(Buyer2Proposal_Query.UnitId) AS Units,
  Buyer2Proposal_Query.ProposalID,
  Buyer2Proposal_Query.Latitude,
  Buyer2Proposal_Query.Longitude
FROM
  Buyer2Proposal_Query
GROUP BY
  Buyer2Proposal_Query.LocationID,
  Buyer2Proposal_Query.SourceCompany,
  Buyer2Proposal_Query.Location,
  Buyer2Proposal_Query.ProposalID,
  Buyer2Proposal_Query.Latitude,
  Buyer2Proposal_Query.Longitude;
