SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SCHNo1,
  Schedule.SchNo,
  Schedule.SchID,
  MstrLease.MLID
FROM
  (
    Clients
    INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
  )
  INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
ORDER BY
  Clients.ClientShNm,
  Schedule.SCHNo1 DESC,
  Schedule.SchID DESC;
