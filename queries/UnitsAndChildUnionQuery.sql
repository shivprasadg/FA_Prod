SELECT
  unitID,
  UnitFinalCost
FROM
  Units
UNION ALL
SELECT
  unitRef,
  ChildFinalCost
FROM
  ChildParts;
