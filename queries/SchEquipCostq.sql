SELECT
  Schedule.SchID,
  Sum(SchGrp.OrigEquipCost) AS SumOfOrigEquipCost
FROM
  Schedule
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
GROUP BY
  Schedule.SchID
ORDER BY
  Sum(SchGrp.OrigEquipCost);
