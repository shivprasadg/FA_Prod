SELECT
  Clients.ClientID,
  Clients.ClientShNm AS ShortName,
  Clients.clientCompanyName AS CompanyName
FROM
  Clients
WHERE
  (
    (
      (Clients.ClientShNm) Not Like "*demo*"
    )
    AND (
      (Clients.clientType)= "Client"
      Or (Clients.clientType)= "Guarantor"
      Or (Clients.clientType) Like "*Target"
    )
  )
ORDER BY
  Clients.ClientShNm;
