SELECT
  Contacts.contactID,
  [clientCompanyName] & Chr(13)& Chr(10)& Nz(
    [ContactAddress1], [clientAddress1]
  )& Chr(13)& Chr(10)& IIf(
    IsNull(
      Nz(
        [ContactAddress2], [clientAddress2]
      )
    ),
    Null,
    Nz(
      [ContactAddress2], [clientAddress1]
    )& Chr(13)& Chr(10)
  )& Nz([ContactCity], [clientCity])& " " & Nz(
    [Contacts].[LState], [Clientstate]
  )& " " & Left(
    Nz([contactZip], [clientzip]),
    5
  ) AS ContactAddress,
  Contacts.contactFullName,
  Contacts.contactEmail,
  Contacts.contactPrimaryPhone AS contactPhone,
  Clients.clientCompanyName,
  Clients.clientID,
  Clients.clientAddress1
FROM
  Clients
  INNER JOIN Contacts ON Clients.clientID = Contacts.clientRef;
