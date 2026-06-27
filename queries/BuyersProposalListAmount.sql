SELECT
  BuyersProposal.ProposalID,
  UCase(
    Left([SalesPerson], 2)
  ) AS SP,
  Buyers.ClientShNm AS Buyer,
  BuyersProposal.ProposalName,
  BuyerProposalAutoStatus.ProposalStatus AS Status,
  Format(
    Nz(
      [LastContactDate], [FirstContactDate]
    ),
    "mm/dd/yy"
  ) AS [Last],
  Count(Rmkt.UnitRef) AS Units,
  Format(
    Sum([TargetSale]),
    "$#,###"
  ) AS Amount,
  BuyersProposal.BuyerID,
  BuyersProposal.SalesPersonID
FROM
  (
    (
      (
        BuyersProposal
        INNER JOIN Buyers ON BuyersProposal.BuyerID = Buyers.clientID
      )
      INNER JOIN BuyersProposalUnits ON BuyersProposal.ProposalID = BuyersProposalUnits.ProposalID
    )
    INNER JOIN Rmkt ON BuyersProposalUnits.UnitID = Rmkt.UnitRef
  )
  INNER JOIN BuyerProposalAutoStatus ON BuyersProposal.ProposalID = BuyerProposalAutoStatus.ProposalID
GROUP BY
  BuyersProposal.ProposalID,
  UCase(
    Left([SalesPerson], 2)
  ),
  Buyers.ClientShNm,
  BuyersProposal.ProposalName,
  BuyerProposalAutoStatus.ProposalStatus,
  Format(
    Nz(
      [LastContactDate], [FirstContactDate]
    ),
    "mm/dd/yy"
  ),
  BuyersProposal.BuyerID,
  BuyersProposal.SalesPersonID;
