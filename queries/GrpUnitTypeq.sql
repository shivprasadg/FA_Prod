SELECT
  SchGrp.SGrpID AS UTGrpRef,
  SchGrp.SGrpID,
  Avg(
    [unitfinalcost] + IIf(
      IsNull([SumOfchildFinalCost]),
      0,
      [sumofchildfinalcost]
    )
  ) AS UnitCost,
  Sum(Units.UnitSynResid) AS SumOfUnitSynResid,
  Avg(Units.UNITRENT) AS AvgOfUNITRENT
FROM
  SchGrp
  INNER JOIN (
    Units
    LEFT JOIN ChildPartSummaryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF
  ) ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  SchGrp.SGrpID,
  SchGrp.SGrpID;
