SELECT
  Salesperson.SPInitials,
  Clients.ClientType,
  Clients.clientCompanyName,
  Max(Notes.noteDateTime) AS MaxOfnoteDateTime
FROM
  Notes
  INNER JOIN (
    Clients
    INNER JOIN Salesperson ON Clients.SalespersonREF = Salesperson.SalespersonID
  ) ON Notes.clientRef = Clients.clientID
GROUP BY
  Salesperson.SPInitials,
  Clients.ClientType,
  Clients.clientCompanyName,
  Notes.CallType
HAVING
  (
    (
      (Clients.ClientType)= "target"
      Or (Clients.ClientType) Like "2nd*"
    )
    AND (
      (Notes.CallType)= "on-site"
    )
  )
ORDER BY
  Salesperson.SPInitials,
  Clients.ClientType DESC;
