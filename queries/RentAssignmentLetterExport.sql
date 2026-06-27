SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  Schedule.SchShortDesc,
  SchGrp.SGrpNO,
  Units.UNITUNITNUM AS [Unit No],
  Units.UNITVIN AS VIN,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] AS Description,
  Units.UNITINSERV AS [In-Service Date],
  [unitfinalcost] + IIf(
    IsNull([sumofchildfinalcost]),
    0,
    [sumofchildfinalcost]
  ) AS [Lessor's Cost(per Unit)],
  1 AS [Item #],
  Right([unitvin], 6) AS vin4,
  schgrpcntq.CountOfSGrpID AS GrpCnt,
  SchGrp.SgrpDesc,
  Schedule.SchID,
  MstrLease.Client AS ClientID
FROM
  (
    (
      MstrLease
      INNER JOIN (
        Schedule
        INNER JOIN schgrpcntq ON Schedule.SchID = schgrpcntq.SchID
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    Units
    LEFT JOIN ChildPartSummaryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF
  ) ON SchGrp.SGrpID = Units.GroupREF;
