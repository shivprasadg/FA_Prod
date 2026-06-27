SELECT
  Deals.dealID AS ID,
  Null AS IsDeleted,
  Deals.clientref AS AccountId,
  0 AS IsPrivate,
  Deals.dealShortDesc AS Name,
  Deals.dealShortDesc AS Description,
  IIf(
    IsNull([dealoutcome]),
    IIf(
      IsNull([deallpafsent]),
      IIf(
        IsNull([dealLpafdue]),
        [transType],
        "Proposal Due"
      ),
      "Proposal Sent"
    ),
    [dealoutcome]
  ) AS StageNm,
  "tbd" AS SortOrder,
  IIf(
    IsNull([dealgrandtotal]),
    0,
    [dealgrandtotal]
  ) AS Amount,
  0 AS Probability,
  0 AS ExpectedRevenue,
  Deals.dealunitqty AS TotalOpportunityQuantity,
  Null AS CloseDate,
  IIf(
    (
      [clienttype] <> "client"
      And [clienttype] <> "guarantor"
    )
    Or IsNull([dealentereddate])
    Or [dealentereddate] < [clientbecamedate],
    "New Business",
    "Existing Business"
  ) AS Type,
  IIf(
    IsNull([dealoutcome]),
    "Awarded",
    Null
  ) AS NextStep,
  Null AS LeadSource,
  IIf(
    IsNull([dealoutcome]),
    0,
    1
  ) AS [Is Closed],
  IIf([dealoutcome] = "awarded", 1, 0) AS [Is Won],
  IIf(
    IsNull([dealoutcome]),
    "Pipeline",
    "Closed"
  ) AS ForecastCategory,
  IIf(
    IsNull([dealoutcome]),
    "Pipeline",
    "Closed"
  ) AS ForecastCategoryName,
  Null AS CampaignId,
  Null AS HasOpportunityLineItem,
  Null AS Pricebook2Id,
  Clients.loinitial AS OwnerId,
  Deals.dealEnteredDate AS Createdate,
  Clients.loinitial AS CreatedById,
  Null AS dealLPAFDueModifiedOnifiedDate,
  Null AS ModifiedOnifiedById,
  Null AS SystemModstamp,
  Null AS LastActivityDate,
  Null AS LastStageChangeDate,
  Clients.creditFYE,
  Null AS FiscalYear,
  Null AS FiscalQuarter,
  Null AS PrimaryPartnerAccountId,
  Null AS Stage_Changed_C,
  Deals.faid AS FAID__c
FROM
  Clients
  INNER JOIN (
    Deals
    LEFT JOIN lookup_transTypes ON Deals.transTypeRef = lookup_transTypes.TransTypeID
  ) ON Clients.clientID = Deals.clientRef
WHERE
  (
    (
      (Deals.dealShortDesc)<> "Tracking Units"
    )
    AND (
      (
        Left([dealshortdesc], 3)
      )<> "dup"
    )
    AND (
      (
        Left([dealshortdesc], 8)
      )<> "new deal"
    )
  );
