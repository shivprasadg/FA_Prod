SELECT
  Schedule.SchMLRef AS MLID,
  Schedule.SchID,
  SchGrp.SGrpID AS GrpID,
  Count(Units.UnitID) AS UnitCount
FROM
  (
    Units
    INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
  )
  INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
GROUP BY
  Schedule.SchMLRef,
  Schedule.SchID,
  SchGrp.SGrpID
ORDER BY
  Schedule.SchID DESC,
  SchGrp.SGrpID DESC;
