SELECT
  Client_with_Units.ClientGroupId,
  Clients.clientID,
  Clients.clientCompanyName,
  Clients.clientAddress1,
  Clients.clientAddress2,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientZip,
  Clients.clientcounty,
  Clients.clientCountry
FROM
  Client_with_Units
  INNER JOIN Clients ON Client_with_Units.clientID = Clients.clientID;
