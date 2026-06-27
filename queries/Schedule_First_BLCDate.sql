SELECT
  MstrLease.Client AS ClientID,
  MstrLease.MLID,
  Nz([SgrpFAID], [SchFAID]) AS FAID,
  Schedule.DateSchAdded,
  SchGrp.SGrpSchID AS ScheduleID,
  Schedule.SchNo AS Schedule,
  SchGrp.Assignee_Bank AS Assignee,
  Min(SchGrp.BLCD) AS FirstBLCD,
  SchGrp.LeaseType
FROM
  (
    Schedule
    INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
GROUP BY
  MstrLease.Client,
  MstrLease.MLID,
  Nz([SgrpFAID], [SchFAID]),
  Schedule.DateSchAdded,
  SchGrp.SGrpSchID,
  Schedule.SchNo,
  SchGrp.Assignee_Bank,
  SchGrp.LeaseType
HAVING
  (
    (
      (SchGrp.LeaseType)= "FMV"
    )
  )
ORDER BY
  SchGrp.SGrpSchID;
