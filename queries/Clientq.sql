SELECT
  Clients.clientCompanyName,
  Clients.clientDBA,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientBecameDate,
  Clients.clientID
FROM
  Clients
WHERE
  (
    (
      (Clients.clientCompanyName)<> "TEST"
    )
    AND (
      (Clients.clientType)= "CLIENT"
    )
  )
ORDER BY
  Clients.clientCompanyName;
