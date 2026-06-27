SELECT
  Clients.clientID,
  Clients.clientCompanyName,
  Clients.ClientShNm,
  Clients.clientType,
  Clients.clientState,
  Clients.clientCity,
  "Edit" AS Edit,
  Salesperson.SPInitials
FROM
  Clients
  LEFT JOIN Salesperson ON Clients.LOinitial = Salesperson.SPInitials
WHERE
  (
    (
      (Clients.clientType)<> "Bank"
      And (Clients.clientType)<> "Vendor"
      And (Clients.clientType)<> "Remarketing"
      And (Clients.clientType)<> "FA"
      And (Clients.clientType)<> "Dead"
    )
  )
ORDER BY
  Clients.clientCompanyName;
