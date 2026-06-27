SELECT
  Clients.ClientShNm,
  Clients.clientWebsite
FROM
  Clients
WHERE
  (
    (
      (Clients.clientWebsite) Is Not Null
    )
  )
ORDER BY
  Clients.ClientShNm;
