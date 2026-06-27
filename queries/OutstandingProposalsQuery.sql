SELECT
  ClientsView_query.clientCompanyName,
  ClientsView_query.clientID,
  Deals.dealShortDesc,
  ClientsView_query.ClientShNm,
  Deals.dealID,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
  Deals.dealLPAFSent,
  Deals.dealUnitQty,
  lookup_transTypes.TransType,
  Deals.dealGrandTotal,
  Groups.groupTotal,
  Groups.groupYear,
  Groups.groupMake,
  Groups.groupModel,
  Groups.groupType,
  Groups.groupUnits,
  Deals.dealLPAFSent,
  Deals.dealoutcomedate,
  Deals.dealoutcome,
  Clients.LOinitial,
  Deals.faid
FROM
  (
    (
      (
        ClientsView_query
        INNER JOIN Deals ON ClientsView_query.clientID = Deals.clientRef
      )
      INNER JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
    )
    INNER JOIN Clients ON ClientsView_query.clientID = Clients.clientID
  )
  INNER JOIN Groups ON Deals.dealID = Groups.dealRef
WHERE
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
    AND (
      (Deals.dealoutcomedate) Is Null
    )
    AND (
      (Groups.AltEquip)= False
    )
    AND (
      (
        ClientsView_query.ClientGroupID
      )<> 18
    )
  );
