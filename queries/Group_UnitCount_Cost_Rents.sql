SELECT
  Units.GroupREF AS GroupID,
  Count(Units.UnitID) AS UnitCount,
  Sum(
    Nz([UnitFinalCost], 0)
  ) AS UnitCost,
  Sum(
    Nz([childFinalCost], 0)
  ) AS ChildPartCost,
  Sum(
    Nz([UNITRENT], 0)
  ) AS SyndRent,
  Sum(
    Nz([UNITSRENT], 0)
  ) AS LORent
FROM
  Units
  INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
GROUP BY
  Units.GroupREF;
