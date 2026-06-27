SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  Min(Units.UnitPDStartDate) AS MinOfUnitPDStartDate,
  Schedule.SchID,
  Min(Units.UnitAcceptDate) AS MinOfUnitAcceptDate,
  Min(Units.unitinserv) AS MinOfunitinserv
FROM
  (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  MstrLease.MLNo,
  Schedule.SchNo,
  Schedule.SchID;
