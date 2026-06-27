SELECT
  ClientsView_query.SPInitials AS Expr1,
  ClientsView_query.clientCompanyName,
  Deals.dealLPAFSent,
  Deals.dealoutcomedate,
  Deals.transTypeRef,
  Deals.dealShortDesc,
  ClientsView_query.ClientShNm,
  Deals.dealID,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
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
  Deals.dealoutcome,
  ClientsView_query.clientshnm
FROM
  (
    (
      ClientsView_query
      INNER JOIN Deals ON ClientsView_query.clientID = Deals.clientRef
    )
    LEFT JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
  )
  INNER JOIN Groups ON Deals.dealID = Groups.dealRef
WHERE
  (
    (
      (Deals.dealoutcomedate) Is Null
    )
    And (
      (lookup_transTypes.TransType)= "Proposal"
    )
    And (
      (Deals.dealLPAFSent) Is Not Null
      And (Deals.dealLPAFSent)<= #3/31/2012#
    )
    And (
      (ClientsView_query.SPInitials)= "sm"
      Or (ClientsView_query.SPInitials)= "hs"
    )
    And (
      (ClientsView_query.clientID)<> 4133
    )
    And (
      (Groups.AltEquip)= No
    )
  )
  Or (
    (
      (Deals.dealoutcomedate) Is Not Null
      And (Deals.dealoutcomedate)> #2/29/2012#
    )
    And (
      (lookup_transTypes.TransType)= "Proposal"
    )
    And (
      (Deals.dealLPAFSent) Is Not Null
      And (Deals.dealLPAFSent)<= #3/31/2012#
    )
    And (
      (ClientsView_query.SPInitials)= "sm"
      Or (ClientsView_query.SPInitials)= "hs"
    )
    And (
      (ClientsView_query.clientID)<> 4133
    )
    And (
      (Groups.AltEquip)= No
    )
  )
ORDER BY
  ClientsView_query.SPInitials;
