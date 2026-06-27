SELECT
  ChildParts.UnitREF,
  Sum(ChildParts.childFinalCost) AS SumOfchildFinalCost,
  ChildParts.ExclFrGrpOEC,
  Units.UnitMake
FROM
  Units
  INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ExclFrGrpOEC,
  Units.UnitMake
HAVING
  (
    (
      (ChildParts.ExclFrGrpOEC)= No
    )
    AND (
      (Units.UnitMake)= "peterbilt"
    )
  );
