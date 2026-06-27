SELECT
  [clientcompanyName] & IIf(
    [clientDBA] <> "", " d/b/a " & [clientDBA],
    ""
  ) AS ClientName,
  *
FROM
  clients
WHERE
  (
    (
      (clients.clientID)<> 4133
    )
    AND (
      (clients.clientType)= "Remarketing"
      Or (clients.clientType) Like "*RMKT*"
    )
  );
