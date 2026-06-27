SELECT
  Salesperson.SalespersonName,
  Deals.transTypeRef,
  Deals.dealOutcome,
  Deals.dealOutcomeDate,
  Deals.dealLPAFSent,
  Deals.dealID,
  Deals.clientRef,
  Deals.dealRFPDate,
  Deals.dealLPAFSent,
  Clients.clientCompanyName,
  Deals.dealUnitQty,
  Groups.groupYear,
  Groups.groupMake,
  Groups.groupModel,
  IIf(
    Left([groups].[groupType], 6)= "daycab",
    "Daycab",
    [groups].[grouptype]
  ) AS Type,
  Deals.dealGrandTotal,
  Groups.groupTotal,
  Deals.dealSchedNum,
  Deals.dealOutcomeReason,
  Groups.groupUnits,
  lookup_transTypes.TransType,
  IIf(
    IsNull([groupunits]),
    [dealunitqty],
    [groupunits]
  ) AS UQTY,
  IIf(
    IsNull([grouptotal]),
    [dealgrandtotal],
    [grouptotal]
  ) AS [D$],
  Deals.FAID,
  Groups.AltEquip,
  Salesperson.SalespersonID,
  Clients.clientBecameDate,
  DateDiff(
    "d",
    [clientbecamedate],
    Date()
  ) AS Age,
  IIf(
    DateDiff(
      "d",
      [clientbecamedate],
      Date()
    )<= 365,
    2,
    1
  ) AS RptGrp,
  Groups.groupType
FROM
  (
    (
      Clients
      INNER JOIN Salesperson ON Clients.SalespersonREF = Salesperson.SalespersonID
    )
    INNER JOIN (
      Deals
      INNER JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
    ) ON Clients.clientID = Deals.clientRef
  )
  LEFT JOIN (
    (
      (
        lookup_GroupMakes
        RIGHT JOIN Groups ON lookup_GroupMakes.groupMakeID = Groups.groupMakeRef
      )
      LEFT JOIN lookup_groupModels ON Groups.groupModelRef = lookup_groupModels.groupModelID
    )
    LEFT JOIN lookup_groupTypes ON Groups.groupTypeRef = lookup_groupTypes.groupTypeID
  ) ON Deals.dealID = Groups.dealRef
WHERE
  (
    (
      (Salesperson.SalespersonName) Like "s*"
      Or (Salesperson.SalespersonName) Like "H*"
    )
    AND (
      (Deals.transTypeRef)<> 1
    )
    AND (
      (Deals.dealOutcome) Is Not Null
      And (
        (Deals.dealOutcome)= "dead"
        Or (Deals.dealOutcome)= "lost"
      )
    )
    AND (
      (Deals.dealOutcomeDate)>= #1/1/2012#
      And (Deals.dealOutcomeDate)<= #3/31/2012#
    )
    AND (
      (Deals.clientRef)<> 4133
    )
    AND (
      (lookup_transTypes.TransType)= "Proposal"
    )
    AND (
      (Groups.AltEquip)= No
    )
  )
ORDER BY
  Salesperson.SalespersonName,
  Deals.dealOutcome,
  Deals.dealID;
