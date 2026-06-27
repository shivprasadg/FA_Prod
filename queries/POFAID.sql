SELECT
  DISTINCT Clients.ClientGroupID,
  Deals.FAID,
  Clients.ClientShNm,
  Deals.dealOutcome
FROM
  Deals
  INNER JOIN Clients ON Deals.clientRef = Clients.clientID
GROUP BY
  Clients.ClientGroupID,
  Deals.FAID,
  Clients.ClientShNm,
  Deals.dealOutcome
HAVING
  (
    (
      (Clients.ClientGroupID)= forms![po selection]!ClientGroupID
    )
    And (
      (Deals.dealOutcome) Is Null
    )
  )
  Or (
    (
      (Clients.ClientGroupID)= forms![po selection]!ClientGroupID
    )
    And (
      (Deals.dealOutcome) Is Not Null
      And (Deals.dealOutcome)<> "dead"
      And (Deals.dealOutcome)<> "lost"
    )
  )
ORDER BY
  Deals.FAID DESC;
