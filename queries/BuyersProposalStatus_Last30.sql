SELECT
  BuyersProposal.SalesPersonID,
  BuyersProposal.SalesPerson,
  Count(BuyersProposal.ProposalID) AS Proposals,
  Sum(
    IIf(
      [ProposalStatus] Not Like '*sold*'
      And [ProposalStatus] <> 'Closed',
      1, 0
    )
  ) AS ProposalActive,
  Sum(
    IIf(
      [ProposalStatus] Like '*sold*', 1,
      0
    )
  ) AS ProposalSold
FROM
  BuyersProposal
WHERE
  (
    (
      (
        BuyersProposal.FirstContactDate
      )>= DateAdd(
        "m",
        -1,
        Date()
      )
    )
  )
GROUP BY
  BuyersProposal.SalesPersonID,
  BuyersProposal.SalesPerson;
