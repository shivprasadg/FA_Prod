SELECT
  Clients.clientID AS [Unique ID],
  Clients.clientCompanyName AS NAME,
  Clients.clienttype AS Type,
  GuarantorsfParent.clientID AS Parent,
  Clients.clientAddress1 AS BILLINGSTREET,
  Clients.clientCity AS BILLINGCITY,
  Clients.clientState AS BILLINGSTATE,
  Clients.clientZip AS BILLINGPOSTALCODE,
  Clients.clientcountry AS BILLINGCOUNTRY,
  Null AS SHIPPINGSTREET,
  Null AS SHIPPINGCITY,
  Null AS SHIPPINGSTATE,
  Null AS SHIPPINGPOSTALCODE,
  Null AS SHIPPINGCOUNTRY,
  Clients.clientMainPhone AS PHONE,
  Clients.clientFax AS Fax,
  Clients.clientWebsite AS Website,
  Clients.clientIndustry AS INDUSTRY,
  0 AS ANNUALREVENUE,
  0 AS NUMBEROFEMPLOYEES,
  Null AS DESCRIPTION,
  Clients.LOinitial AS OWNER,
  Clients.fleetDaycabs AS DAYCABS__C,
  Clients.fleetSleepers AS SLEEPERS__C,
  Clients.fleetsize AS FLEET_SIZE__C,
  Clients.creditMoodyRating AS CREDIT_RATING_MOODYS_SP__C,
  Clients.creditRank AS FA_CREDIT_RANK__C,
  Clients.creditFAScore AS FA_CREDIT_SCORE__C,
  Clients.creditFinance AS ACQUISITION__C,
  Clients.creditmaint AS MAINTENANCE__C,
  Clients.FleetReplacementCycle AS CYCLE_YEARS__C,
  Clients.clientOBC AS OBC__C,
  Clients.ClientGroupID
FROM
  Clients
  LEFT JOIN GuarantorsfParent ON Clients.ClientGroupID = GuarantorsfParent.ClientGroupID
WHERE
  (
    (
      (Clients.clienttype)<> "Bank"
      And (Clients.clienttype)<> "dead"
      And (Clients.clienttype)<> "r*"
    )
  )
ORDER BY
  Clients.clientID,
  Clients.clientCompanyName;
