SELECT
  ClientsView_query.clientID,
  Deals.transTypeRef,
  Deals.dealShortDesc,
  ClientsView_query.ClientShNm,
  Deals.dealID,
  ClientsView_query.CLIENTGRPMAX,
  ClientsView_query.clientCompanyName,
  Deals.dealLPAFSent,
  Deals.dealUnitQty,
  Deals.dealOutcome,
  Deals.dealOutcomeDate,
  lookup_transTypes.TransType,
  Deals.dealGrandTotal,
  Groups.groupTotal,
  Groups.groupYear,
  Groups.groupMake,
  Groups.groupModel,
  Groups.groupType,
  Groups.groupUnits,
  Groups.AltEquip,
  Salesperson.SPInitials
FROM
  (
    (
      (
        ClientsView_query
        INNER JOIN Deals ON ClientsView_query.clientID = Deals.clientRef
      )
      INNER JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
    )
    INNER JOIN Groups ON Deals.dealID = Groups.dealRef
  )
  INNER JOIN Salesperson ON ClientsView_query.SalespersonREF = Salesperson.SalespersonID
WHERE
  (
    (
      (ClientsView_query.clientID)<> 4133
    )
    And (
      (Deals.dealLPAFSent) Is Not Null
      And (Deals.dealLPAFSent)<= forms!LorPtm!OPAsof
    )
    And (
      (Deals.dealOutcome) Is Null
    )
    And (
      (lookup_transTypes.TransType)= "Proposal"
    )
    And (
      (Groups.AltEquip)= No
    )
  )
  Or (
    (
      (ClientsView_query.clientID)<> 4133
    )
    And (
      (Deals.dealLPAFSent) Is Not Null
      And (Deals.dealLPAFSent)<= forms!LorPtm!OPAsof
    )
    And (
      (Deals.dealOutcome) Is Not Null
    )
    And (
      (Deals.dealOutcomeDate) Is Not Null
      And (Deals.dealOutcomeDate)> forms!LorPtm!OPAsof
    )
    And (
      (lookup_transTypes.TransType)<> "Opportunity"
    )
    And (
      (Groups.AltEquip)= No
    )
  )
  Or (
    (
      (Deals.dealLPAFSent)>+ #1/1/2012#
      And (Deals.dealLPAFSent)<= #2/29/2012#
    )
  );
