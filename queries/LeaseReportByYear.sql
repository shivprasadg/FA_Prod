SELECT
  LeaseReport2.LeaseYear,
  LeaseReport2.Lessee,
  LeaseReport2.[Lease No],
  LeaseReport2.Make,
  LeaseReport2.UnitType,
  Count(LeaseReport2.AssetId) AS UnitCount,
  Sum(LeaseReport2.UnitCost) AS UnitCost,
  Sum(LeaseReport2.Refrig) AS Refrig,
  Sum(
    [MiscCost] + [CoNameCost] + [RegDecalCost]
  ) AS Decals,
  Sum(LeaseReport2.OBC) AS OBC,
  Sum(
    [LeaseReport2].[OBCInstall] + [OBCAdapt] + [OBCDome] + [OBCMulti]
  ) AS OBCInstall,
  Sum(LeaseReport2.FeesTax) AS FeesTax,
  Sum(LeaseReport2.Freight) AS Freight,
  Sum(LeaseReport2.Warranty) AS Warranty,
  Sum(LeaseReport2.Inspection) AS Inspection,
  Sum(LeaseReport2.OilRig) AS OilRig,
  Sum(
    ([TankBody] + [TankOther])
  ) AS TankBodies,
  Sum(LeaseReport2.PTO) AS PTO,
  Sum(LeaseReport2.Other) AS OtherCost,
  Sum(LeaseReport2.TotalChildCost) AS TotalChildCost,
  Sum(LeaseReport2.TotalCost) AS TotalCost
FROM
  LeaseReport2
GROUP BY
  LeaseReport2.LeaseYear,
  LeaseReport2.Lessee,
  LeaseReport2.[Lease No],
  LeaseReport2.Make,
  LeaseReport2.UnitType;
