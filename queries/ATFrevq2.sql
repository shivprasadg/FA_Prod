SELECT
  Clients.clientgroupID,
  Clients.ClientShNm,
  Schedule.SchNo,
  SchGrp.BLCD,
  AtfRevq1.RevNo,
  AtfRevq1.RevDt,
  AtfRevq1.ReviewDt,
  AtfRevq1.reviewer
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN (
      Schedule
      INNER JOIN AtfRevq1 ON Schedule.SchID = AtfRevq1.schID
    ) ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (SchGrp.BLCD) Is Null
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
  )
ORDER BY
  Clients.clientgroupID,
  Clients.ClientShNm,
  Schedule.SchID;
