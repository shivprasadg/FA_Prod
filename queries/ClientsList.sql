SELECT
  Clients.clientCompanyName AS Client,
  Clients.clientID,
  Clients.ClientShNm
FROM
  Clients
WHERE
  (
    (
      (Clients.clientType)<> "VENDOR"
      And (Clients.clientType)<> "BANK"
      And (Clients.clientType)<> "dead"
      And (Clients.clientType)<> "Remarketing"
    )
  );
