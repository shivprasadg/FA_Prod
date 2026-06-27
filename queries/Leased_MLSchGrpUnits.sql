SELECT
  DISTINCT Schedule.SchID,
  Schedule.SchNo,
  MstrLease.MLID,
  Count(Units.UnitID) AS Units,
  MstrLease.Client,
  MstrLease.MLNo,
  MstrLease.MLOrig
FROM
  (
    MstrLease
    INNER JOIN (
      (
        Schedule
        INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
      )
      INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
    ) ON MstrLease.MLID = Schedule.SchMLRef
  )
  LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
WHERE
  (
    (
      (Units.UnitStatus)<> "A"
      And (Units.UnitStatus)<> "B"
    )
  )
GROUP BY
  Schedule.SchID,
  Schedule.SchNo,
  MstrLease.MLID,
  MstrLease.Client,
  MstrLease.MLNo,
  MstrLease.MLOrig
ORDER BY
  Schedule.SchNo,
  MstrLease.MLID;
