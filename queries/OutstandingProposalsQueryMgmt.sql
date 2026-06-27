SELECT
  DISTINCT ClientsView_query.ClientShNm,
  ClientsView_query.clientID,
  Clients.LOinitial,
  Deals.transTypeRef,
  Deals.dealShortDesc,
  ClientsView_query.ClientShNm,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
  Deals.dealLPAFSent,
  Deals.dealUnitQty,
  lookup_transTypes.TransType,
  Deals.dealGrandTotal,
  Deals.dealLPAFSent,
  Deals.dealoutcome,
  Deals.dealID,
  Deals.faid
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
  Deals.dealShortDesc,
  ClientsView_query.ClientShNm,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
  Deals.dealUnitQty,
  lookup_transTypes.TransType,
  Deals.dealGrandTotal,
  Deals.dealLPAFSent,
  Deals.dealoutcome,
  Deals.dealID,
  Deals.faid,
  ClientsView_query.ClientShNm,
  Deals.dealLPAFSent
HAVING
  (
    (
      (ClientsView_query.clientID)<> 4133
    )
    AND (
      (lookup_transTypes.TransType)= "Proposal"
    )
    AND (
      (Deals.dealLPAFSent) Is Not Null
    )
  );
