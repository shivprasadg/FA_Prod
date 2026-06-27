SELECT
  Clients.clientID,
  Clients.ClientShNm,
  Clients.clientCompanyName,
  MstrLease.MLOrig,
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
      (MstrLease.MLOrig)= "fa"
    )
  );
