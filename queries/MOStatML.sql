SELECT
  MstrLease.MLSent,
  Clients.SalespersonREF,
  MstrLease.MLNo,
  MstrLease.MLBack,
  Clients.clienttype,
  Clients.clientCompanyName,
  Clients.clientBecameDate
FROM
  Clients
  INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
WHERE
  (
    (
      (MstrLease.MLNo)>= "Enter start dt]"
      And (MstrLease.MLNo)<= [enter end dt]
    )
  )
  OR (
    (
      (MstrLease.MLBack)>= "Enter start dt]"
      And (MstrLease.MLBack)<= [enter end dt]
    )
  )
ORDER BY
  MstrLease.MLSent,
  Clients.SalespersonREF,
  MstrLease.MLNo;
