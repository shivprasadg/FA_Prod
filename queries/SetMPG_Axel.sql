UPDATE
  (
    (
      (
        Units
        INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
      )
      INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
    )
    INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID
  )
  INNER JOIN Clients ON MstrLease.Client = Clients.clientID
SET
  Units.MpgGoal = 6,
  Units.DrivingMpgGoal = 6
WHERE
  (
    (
      (Units.clientgroupid)= 151
    )
  );
