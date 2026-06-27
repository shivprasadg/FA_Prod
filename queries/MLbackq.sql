SELECT
  Clients.clientID,
  MstrLease.MLID,
  MstrLease.MLOrig,
  MstrLease.MLBack
FROM
  Clients
  INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
WHERE
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Clients.clienttype)= "client"
    )
  )
ORDER BY
  MstrLease.MLID;
