SELECT
  ChildParts.UnitREF,
  Sum(ChildParts.childFinalCost) AS SumOfchildFinalCost,
  ChildParts.ChildType
FROM
  ChildParts
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildType;
