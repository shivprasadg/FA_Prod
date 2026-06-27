SELECT
  Clients_FAID.FAID,
  First(Clients_FAID_1.ClientShNm) AS ClientShNm,
  First(Clients_FAID_1.clientID) AS clientID
FROM
  Clients_FAID
  INNER JOIN Clients_FAID AS Clients_FAID_1 ON Clients_FAID.FAID = Clients_FAID_1.FAID
GROUP BY
  Clients_FAID.FAID
ORDER BY
  Clients_FAID.FAID,
  First(Clients_FAID_1.ClientShNm);
