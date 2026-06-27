SELECT
  Clients.clientIndustry,
  Clients.ClientGroupID,
  Clients.clientCompanyName,
  Clients.clientType,
  Clients.creditFAProfile,
  Clients.creditFAProfile,
  Clients.creditFAScore,
  Clients.creditFAScore,
  Clients.cpadded
FROM
  Clients
WHERE
  (
    (
      (Clients.ClientGroupID)<> 18
    )
    AND (
      (Clients.clientType)= "client"
      Or (Clients.clientType) Like "*target"
    )
  )
ORDER BY
  Clients.clientIndustry,
  Clients.ClientGroupID,
  Clients.clientCompanyName;
