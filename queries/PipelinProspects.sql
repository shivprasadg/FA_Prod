INSERT INTO FAScRptT1 (
  ClientShNm, clientCompanyName, creditFAProfile,
  ClientGroupID, creditFAScore, SPInitials
)
SELECT
  DISTINCT Clients.ClientShNm,
  Clients.clientCompanyName,
  Clients.creditFAProfile,
  Clients.ClientGroupID,
  Clients.creditFAScore,
  Salesperson.SPInitials
FROM
  (
    (
      Clients
      INNER JOIN Salesperson ON Clients.SalespersonREF = Salesperson.SalespersonID
    )
    INNER JOIN Deals ON Clients.clientID = Deals.clientRef
  )
  INNER JOIN Groups ON Deals.dealID = Groups.dealRef
WHERE
  (
    (
      (Clients.ClientShNm)<> "test"
    )
    AND (
      (Deals.dealOutcome) Is Null
      And (Deals.dealOutcome) Is Null
    )
    AND (
      (Deals.transTypeRef)= 2
    )
  )
  OR (
    (
      (Clients.ClientShNm)<> "test"
    )
    AND (
      (Deals.dealOutcome) Is Null
      And (Deals.dealOutcome) Is Null
    )
    AND (
      (Deals.transTypeRef)= 3
    )
  );
