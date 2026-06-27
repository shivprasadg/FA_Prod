SELECT
  clients.clientID,
  [clientcompanyName] & IIf(
    [clientDBA] <> "", " d/b/a " & [clientDBA],
    ""
  ) AS CompanyAndDBA
FROM
  clients;
