SELECT
  Groups.dealRef,
  Sum(
    UnitsAndChildUnionQuery.UnitFinalCost
  ) AS SumOfUnitFinalCost
FROM
  UnitsAndChildUnionQuery
  INNER JOIN (
    UnitsSummary
    INNER JOIN Groups ON UnitsSummary.GroupREF = Groups.groupID
  ) ON UnitsAndChildUnionQuery.unitID = UnitsSummary.UnitID
GROUP BY
  Groups.dealRef;
