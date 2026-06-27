SELECT
  SourceCompany,
  BankPartner AS Bank,
  [Sch|Grp] AS [Lease Sch|Grp],
  UnitId AS AssetID,
  Unitnum,
  UnitVin AS VIN,
  UnitStatus AS [Unit Status],
  UnitYr AS Myear,
  UnitMake AS Make,
  UnitModel AS Model,
  UnitType AS Type,
  UnitSubType AS SubType,
  Axle,
  ParkedLocation AS [Parked Location],
  PState,
  ProjEOL AS [Projected Offlease],
  OffLeaseDate AS [Actual OffLease],
  RStatus AS InvStatus,
  DaysInInventory AS [Days IN],
  Residual AS RV,
  CurrBook,
  TargetSalePrice AS [Target Price],
  [Expected GP],
  TotalExpense AS [RMKT Exp],
  Odometer AS ODO,
  Milesallowed AS [Miles Allowed],
  OverMiles AS [Over|Under Miles],
  OverMilesCost AS [Over Miles Cost],
  Disposition AS Disp,
  BuyOut,
  ClientID,
  BankID
FROM
  vw_RemarketingReportAllUnits
WHERE
  (
    (
      (UnitStatus) IN ('A', 'TR', 'R', 'P')
    )
    AND (
      (BankID)<> 4321
    )
  )
  OR (
    (
      (UnitStatus) LIKE '*E*'
    )
    AND (
      (BankID)<> 4321
    )
  )
ORDER BY
  SourceCompany,
  BankPartner;
