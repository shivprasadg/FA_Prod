SELECT
  RemarketingReportAllUnits.SourceCompany AS [Account Name],
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.Unitnum AS [Cust Unit#],
  RemarketingReportAllUnits.UnitStatus,
  RemarketingReportAllUnits.UnitYr,
  RemarketingReportAllUnits.UnitMake,
  RemarketingReportAllUnits.UnitModel,
  RemarketingReportAllUnits.UnitType,
  vw_SixKeys.Locationsn AS Location,
  "" AS ProjectedParkedDate,
  RemarketingReportAllUnits.DaysInInventory,
  RemarketingReportAllUnits.Odometer,
  [BuyOut] + [TotalExpense] AS TotalCost,
  RemarketingReportAllUnits.UnitSalePrice AS [Actual Sale Price],
  IIf(
    [UnitSalePrice] > 0, [UnitProfit],
    0
  ) AS [PnL on Actual],
  RemarketingReportAllUnits.FMVBase AS FMVBasePrice,
  IIf(
    [FMVBasePrice] > 0, [FMVBase] - [TotalCost],
    0
  ) AS [PnL on FMVBase],
  RemarketingReportAllUnits.FMV,
  IIf([FMV] > 0, [FMV] - [TotalCost], 0) AS [PnL on FMV],
  RemarketingReportAllUnits.TargetSalePrice,
  IIf(
    [TargetSalePrice] > 0, [TargetSalePrice] - [TotalCost],
    0
  ) AS [PnL on Target],
  IIf(
    [FMVBase] > 0, [FMVBase] - [BuyOut],
    0
  ) AS [Term Penalty],
  RemarketingReportAllUnits.BankPartner,
  RemarketingReportAllUnits.BankShareYN,
  RemarketingReportAllUnits.BankShare,
  vw_SixKeys.clientID,
  RemarketingReportAllUnits.BuyerID,
  RemarketingReportAllUnits.UnitID,
  RemarketingReportAllUnits.BuyOutDate,
  RemarketingReportAllUnits.InvoiceSent,
  RemarketingReportAllUnits.InvoicePaid,
  RemarketingReportAllUnits.DaysInInventoryPerAccounting AS [Days Since BuyOut]
FROM
  RemarketingReportAllUnits
  INNER JOIN vw_SixKeys ON RemarketingReportAllUnits.UnitID = vw_SixKeys.AssetID
ORDER BY
  RemarketingReportAllUnits.SourceCompany;
