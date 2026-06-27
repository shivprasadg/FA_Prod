SELECT
  Clients.clientID,
  Clients.ClientShNm,
  Salesperson.SPInitials,
  Clients.clientType,
  [contactfirst] & " " & [contactlast] AS Nm,
  Contacts.contactTitle,
  Contacts.contactRTitle,
  Contacts.contactPrimaryPhone,
  Contacts.contactPhone2,
  Contacts.contactCell,
  Contacts.contactEmail
FROM
  (
    Clients
    LEFT JOIN Contacts ON Clients.clientID = Contacts.clientRef
  )
  INNER JOIN Salesperson ON Clients.SalespersonREF = Salesperson.SalespersonID
WHERE
  (
    (
      (Clients.clientType) Like "*target*"
    )
    AND (
      (Contacts.contactRTitle)= "4"
      Or (Contacts.contactRTitle)= "5"
    )
    AND (
      (Contacts.ContactObsolete)= No
    )
  )
ORDER BY
  Clients.ClientShNm,
  Contacts.contactLast,
  Contacts.contactFirst;
