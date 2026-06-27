SELECT
  Schedule.SchFAID AS FAID,
  Schedule.SchNo AS Schedule,
  SchGrp.SGrpNO AS [Group],
  Count(Units.UnitID) AS Units,
  SchGrp.SGrpID,
  Schedule.SchID
FROM
  (
    Schedule
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  Schedule.SchFAID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.SGrpID,
  Schedule.SchID
HAVING
  (
    (
      (Schedule.SchFAID)<> ""
    )
  )
ORDER BY
  Schedule.SchNo,
  SchGrp.SGrpNO;
