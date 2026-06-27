SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SCHNo1,
  Schedule.SchNo
FROM
  (
    Clients
    INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
  )
  INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
WHERE
  (
    (
      (MstrLease.MLID)= [Forms]![schedules_sub]![MLID]
    )
  )
ORDER BY
  Clients.ClientShNm,
  Schedule.SCHNo1 DESC,
  Schedule.SchID DESC;
