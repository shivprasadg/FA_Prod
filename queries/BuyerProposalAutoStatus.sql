SELECT
  BuyersProposalsStatus.BuyerID,
  BuyersProposalsStatus.ProposalID,
  Count(BuyersProposalsStatus.UnitID) AS Units,
  Sum(
    BuyersProposalsStatus.UnitsSoldCount
  ) AS Sold,
  Sum(
    BuyersProposalsStatus.UnitsNotSold
  ) AS NotSold,
  Sum(
    BuyersProposalsStatus.SelfSoldCount
  ) AS Isold,
  Switch(
    [NotSold] > 0, "Active", [Isold] > 0
    And [Sold] = [Units],
    "Sold", [Sold] = [Units], "Closed"
  ) AS ProposalStatus,
  Max(
    Nz(
      [LastContactDate], [FirstContactDate]
    )
  ) AS LastAction
FROM
  BuyersProposalsStatus
GROUP BY
  BuyersProposalsStatus.BuyerID,
  BuyersProposalsStatus.ProposalID;
