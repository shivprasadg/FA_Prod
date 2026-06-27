SELECT
  BuyersProposal.ProposalID,
  UCase(
    Left([SalesPerson], 2)
  ) AS SP,
  Buyers.ClientShNm AS Buyer,
  BuyersProposal.ProposalName,
  BuyersProposal.ProposalStatus AS Status,
  Nz(
    [LastContactDate], [FirstContactDate]
  ) AS [Last Contact],
  BuyersProposal.BuyerID,
  Count(BuyersProposalUnits.UnitID) AS Units,
  Sum(Rmkt.TargetSale) AS Target
FROM
  (
    (
      BuyersProposal
      INNER JOIN Buyers ON BuyersProposal.BuyerID = Buyers.clientID
    )
    INNER JOIN BuyersProposalUnits ON BuyersProposal.ProposalID = BuyersProposalUnits.ProposalID
  )
  INNER JOIN Rmkt ON BuyersProposalUnits.UnitID = Rmkt.UnitRef
GROUP BY
  BuyersProposal.ProposalID,
  UCase(
    Left([SalesPerson], 2)
  ),
  Buyers.ClientShNm,
  BuyersProposal.ProposalName,
  BuyersProposal.ProposalStatus,
  Nz(
    [LastContactDate], [FirstContactDate]
  ),
  BuyersProposal.BuyerID;
