SELECT
  Clients.creditRank,
  Clients.creditFAProfile,
  Clients.creditFAScore,
  Clients.clientType,
  Clients.clientID
FROM
  Clients
WHERE
  (
    (
      (Clients.clientID)= [forms]![mgmtmenu]![SelCP]
    )
  );
