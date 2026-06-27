SELECT
  Clients.ClientGroupId,
  Clients.clientID,
  Clients.ClientShNm,
  Schedule.SchFAID AS FAID,
  MstrLease.MLID
FROM
  (
    Clients
    INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
  )
  INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
GROUP BY
  Clients.ClientGroupId,
  Clients.clientID,
  Clients.ClientShNm,
  Schedule.SchFAID,
  MstrLease.MLID
HAVING
  (
    (
      (Schedule.SchFAID)<> ""
    )
  )
ORDER BY
  Clients.ClientShNm;
