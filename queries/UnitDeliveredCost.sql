SELECT
  Units.GroupREF AS GroupId,
  Units.UnitID AS AssetId,
  Units.UnitRent AS SyndRent,
  Units.UnitSRent AS LORent,
  Units.unitfinalcost AS UnitCost,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS ChildCost,
  CCur(
    [unitfinalcost] + Sum(
      Nz([childFinalCost], 0)
    )
  ) AS DeliveredCost
FROM
  Units
  LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
GROUP BY
  Units.GroupREF,
  Units.UnitID,
  Units.UnitRent,
  Units.UnitSRent,
  Units.unitfinalcost;
