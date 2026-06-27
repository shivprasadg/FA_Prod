SELECT
  Clients.ClientShNm,
  Clients.clientID,
  MstrLease.MLID,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  MstrLease.MLDt,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.*,
  Schedule.SchBnkClNO,
  Schedule.BLCDACDT,
  Clients.ClientGroupId AS ClientGroupID
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
ORDER BY
  Clients.ClientShNm,
  Schedule.SchID,
  SchGrp.SGrpNO;
