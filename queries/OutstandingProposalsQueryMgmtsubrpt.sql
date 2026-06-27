SELECT
  DISTINCT ClientsView_query.ClientShNm,
  ClientsView_query.clientID,
  Clients.LOinitial,
  Deals.transTypeRef,
  ClientsView_query.ClientShNm,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
  Min(Deals.dealLPAFSent) AS MinOfdealLPAFSent
FROM
  (
    (
      ClientsView_query
      INNER JOIN Deals ON ClientsView_query.clientID = Deals.clientRef
    )
    INNER JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
  )
  INNER JOIN Clients ON ClientsView_query.clientID = Clients.clientID
WHERE
  (
    (
      (Deals.dealLPAFSent) Is Not Null
    )
    AND (
      (lookup_transTypes.TransType)= "Proposal"
    )
    AND (
      (Deals.dealoutcomedate) Is Null
    )
    AND (
      (ClientsView_query.clienttype)<> "client"
    )
    AND (
      (Clients.ClientShNm) Not Like "*demo*"
    )
  )
GROUP BY
  ClientsView_query.clientID,
  Clients.LOinitial,
  Deals.transTypeRef,
  ClientsView_query.ClientShNm,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
  ClientsView_query.ClientShNm
HAVING
  (
    (
      (ClientsView_query.clientID)<> 4133
    )
  );
