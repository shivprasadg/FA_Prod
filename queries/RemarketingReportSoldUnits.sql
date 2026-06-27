SELECT
  RemarketingReportAllUnits.SourceCompany AS [Account Name],
  RemarketingReportAllUnits.Schedule,
  RemarketingReportAllUnits.UnitGroup,
  RemarketingReportAllUnits.UnitNum AS [Unit#],
  RemarketingReportAllUnits.UnitStatus,
  RemarketingReportAllUnits.UnitYr,
  RemarketingReportAllUnits.UnitMake,
  RemarketingReportAllUnits.UnitModel,
  RemarketingReportAllUnits.UnitType,
  RemarketingReportAllUnits.Odometer,
  RemarketingReportAllUnits.Buyer,
  RemarketingReportAllUnits.BuyersOrder,
  RemarketingReportAllUnits.DaysInInventory,
  RemarketingReportAllUnits.Residual,
  RemarketingReportAllUnits.TermPenalty AS [Term Penalty],
  [BuyOut] + Nz([TermPenalty], 0) AS BookValue,
  RemarketingReportAllUnits.BuyOut,
  RemarketingReportAllUnits.TotalExpense AS [RMKT Expense],
  (
    [BuyOut] + Nz([TermPenalty], 0)
  )+ [TotalExpense] AS TotalCost,
  RemarketingReportAllUnits.UnitSalePrice AS [Sale Price],
  [UnitSalePrice] -(
    (
      [BuyOut] + Nz([TermPenalty], 0)
    )+ [TotalExpense]
  ) AS [PnL on Sale],
  RemarketingReportAllUnits.OEC,
  RemarketingReportAllUnits.InvoiceSent,
  RemarketingReportAllUnits.InvoicePaid,
  RemarketingReportAllUnits.InvoiceNum,
  RemarketingReportAllUnits.SalesPerson,
  RemarketingReportAllUnits.UnitVIN,
  RemarketingReportAllUnits.UnitID,
  RemarketingReportAllUnits.DaysInInventoryPerAccounting AS [Days Since BuyOut],
  RemarketingReportAllUnits.BuyOutDate,
  RemarketingReportAllUnits.ClientID,
  RemarketingReportAllUnits.BuyerID
FROM
  RemarketingReportAllUnits
WHERE
  (
    (
      (
        RemarketingReportAllUnits.InvoicePaid
      ) Is Not Null
    )
  )
ORDER BY
  RemarketingReportAllUnits.SourceCompany;
