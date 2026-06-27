SELECT
  UnitsAndChildUnionQuery.unitID,
  Sum(
    UnitsAndChildUnionQuery.UnitFinalCost
  ) AS SumOfUnitFinalCost,
  Groups.groupID,
  Groups.dealRef
FROM
  UnitsAndChildUnionQuery
  INNER JOIN (
    UnitsSummary
    INNER JOIN Groups ON UnitsSummary.GroupREF = Groups.groupID
  ) ON UnitsAndChildUnionQuery.unitID = UnitsSummary.UnitID
GROUP BY
  UnitsAndChildUnionQuery.unitID,
  Groups.groupID,
  Groups.dealRef;
