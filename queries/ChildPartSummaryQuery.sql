SELECT
  SchGrp.SGrpID,
  ChildParts.UnitREF,
  Sum(ChildParts.childFinalCost) AS SumOfchildFinalCost,
  ChildParts.ExclFrGrpOEC,
  Sum(ChildParts.ChildDeposit) AS SumOfChildDeposit
FROM
  (
    SchGrp
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
GROUP BY
  SchGrp.SGrpID,
  ChildParts.UnitREF,
  ChildParts.ExclFrGrpOEC
HAVING
  (
    (
      (ChildParts.ExclFrGrpOEC)= False
    )
  );
