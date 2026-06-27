SELECT
  SchGrp.SGrpSchID AS ScheduleID,
  UnitTotalOEC.GroupID,
  UnitTotalOEC.AssetID,
  IIf(
    IsDate([BLCD])
    And IsDate([AssignmentDate]),
    DateDiff("m", [BLCD], [AssignmentDate]),
    0
  ) AS StripsNum,
  UnitTotalOEC.SalesRent,
  SchGrp.SgrpSynRebate AS RebateUnitActual,
  SchGrp.RebateUnitProjected,
  CCur(
    Nz([SgrpSynRebate], 0)
  ) AS UTA
FROM
  SchGrp
  INNER JOIN UnitTotalOEC ON SchGrp.SGrpID = UnitTotalOEC.GroupID
GROUP BY
  SchGrp.SGrpSchID,
  UnitTotalOEC.GroupID,
  UnitTotalOEC.AssetID,
  IIf(
    IsDate([BLCD])
    And IsDate([AssignmentDate]),
    DateDiff("m", [BLCD], [AssignmentDate]),
    0
  ),
  UnitTotalOEC.SalesRent,
  SchGrp.SgrpSynRebate,
  SchGrp.RebateUnitProjected,
  CCur(
    Nz([SgrpSynRebate], 0)
  );
