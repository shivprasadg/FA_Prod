SELECT
  Schedule.SchID,
  Schedule.SchMLRef AS MLID,
  Schedule.SchNo AS Schedule,
  Count(Units.UnitID) AS Units
FROM
  (
    SchGrp
    INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  Schedule.SchID,
  Schedule.SchMLRef,
  Schedule.SchNo
ORDER BY
  Schedule.SchID,
  Schedule.SchNo;
