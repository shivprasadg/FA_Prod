SELECT
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitVIN
FROM
  (
    (
      MstrLease
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  LEFT JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (MstrLease.MLNo) Like "hbc*"
    )
  )
ORDER BY
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO;
