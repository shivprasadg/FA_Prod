SELECT
  ClientsView_query.clientCompanyName,
  ClientsView_query.clientID,
  ClientsView_query.SPInitials AS Expr1,
  Deals.FAID,
  Deals.transTypeRef,
  Deals.dealShortDesc,
  Deals.dealLPAFSent,
  ClientsView_query.ClientShNm,
  Deals.dealID,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
  DatePart("yyyy", [DEALLPAFSENT]) AS yr,
  DatePart("m", [DEALLPAFSENT]) AS MO,
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
  Deals.FAID
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
      (ClientsView_query.clientID)<> 4133
    )
    And (
      (lookup_transTypes.TransType)= "Proposal"
    )
    And (
      (Deals.dealLPAFSent) Is Not Null
      And (Deals.dealLPAFSent)>= #2/1/2012#
      And (Deals.dealLPAFSent)<= #2/29/2012#
    )
    And (
      (ClientsView_query.SPInitials)= "hs"
      Or (ClientsView_query.SPInitials)= "sm"
    )
    And (
      (Groups.AltEquip)= No
    )
  );
