SELECT
  Clients.ClientShNm,
  Clients.clientID,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpID AS GroupID,
  SchGrp.*,
  Schedule.SchBnkClNO,
  Schedule.BLCDACDT,
  MstrLease.MLOrig,
  MstrLease.MLDt,
  Clients.ClientGroupID,
  MstrLease.MLOrig
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
  MstrLease.MLNo,
  Schedule.SchID,
  SchGrp.SGrpNO;
