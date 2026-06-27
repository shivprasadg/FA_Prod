SELECT
  BuyersProposalUnitCountDistinct.SalesPersonID,
  Count(
    BuyersProposalUnitCountDistinct.UnitID
  ) AS UnitCount
FROM
  BuyersProposalUnitCountDistinct
GROUP BY
  BuyersProposalUnitCountDistinct.SalesPersonID;
