SELECT
  ChildParts.UnitREF,
  Sum(ChildParts.childFinalCost) AS ChildPartCost,
  Sum(ChildParts.ChildDeposit) AS ChildDeposit
FROM
  ChildParts
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ExclFrGrpOEC
HAVING
  (
    (
      (ChildParts.ExclFrGrpOEC)= False
    )
  )
ORDER BY
  ChildParts.UnitREF;
