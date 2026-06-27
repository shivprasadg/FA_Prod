SELECT
  DatePart("yyyy", [dealoutcomedate]) AS [Client Yr],
  Clients.clientCompanyName,
  Clients.clientBecameDate,
  Deals.dealOutcome
FROM
  (
    Clients
    INNER JOIN Deals ON Clients.clientID = Deals.clientRef
  )
  INNER JOIN Documentation ON Deals.dealID = Documentation.dealRef
GROUP BY
  DatePart("yyyy", [dealoutcomedate]),
  Clients.clientCompanyName,
  Clients.clientBecameDate,
  Deals.dealOutcome
HAVING
  (
    (
      (Clients.clientBecameDate) Is Not Null
      And (Clients.clientBecameDate)< #1/1/2010#
    )
    AND (
      (Deals.dealOutcome)= "Awarded"
    )
  )
ORDER BY
  DatePart("yyyy", [dealoutcomedate]),
  Clients.clientCompanyName;
