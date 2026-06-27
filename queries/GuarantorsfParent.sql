SELECT
  Clients.clientID,
  Clients.clientType,
  Clients.ClientShNm,
  Clients.ClientGroupID
FROM
  Clients
WHERE
  (
    (
      (Clients.clientType)= "guarantor"
    )
  )
ORDER BY
  Clients.clientID;
