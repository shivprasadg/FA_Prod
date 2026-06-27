SELECT
  Clients.ClientShNm,
  MstrLease.Client,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitID
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Schedule.SchNo)= "3-savage"
    )
  );
