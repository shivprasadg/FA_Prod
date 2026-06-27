SELECT
  *
FROM
  vw_ScheduleUnitDoc
  INNER JOIN Schedule_Group_WorkTable ON vw_ScheduleUnitDoc.SchID = Schedule_Group_WorkTable.ScheduleID;
