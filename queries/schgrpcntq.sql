SELECT
  Schedule.SchID,
  Count(SchGrp.SGrpID) AS CountOfSGrpID
FROM
  Schedule
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
GROUP BY
  Schedule.SchID;
