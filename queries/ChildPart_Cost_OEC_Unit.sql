SELECT
  ChildParts.UnitREF AS UnitId,
  CCur(
    Nz([childFinalCost], 0)
  ) AS ChildPartCost
FROM
  ChildParts;
