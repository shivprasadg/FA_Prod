SELECT
  BuyerProposalAutoStatus.BuyerID,
  Max(
    BuyerProposalAutoStatus.LastAction
  ) AS LastProposalDate
FROM
  BuyerProposalAutoStatus
GROUP BY
  BuyerProposalAutoStatus.BuyerID;
