SELECT
  Clients.creditFAScore,
  Clients.ClientGroupID,
  Clients.ClientShNm,
  Clients.clientCompanyName,
  Clients.creditFAProfile
FROM
  Clients
WHERE
  (
    (
      (Clients.creditFAScore)>= 0
    )
  )
ORDER BY
  Clients.creditFAScore,
  Clients.ClientGroupID,
  Clients.ClientShNm;
