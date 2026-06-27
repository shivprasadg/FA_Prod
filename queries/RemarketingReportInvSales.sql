SELECT
  RemarketingReportAllUnits.SourceCompany AS [Account Name],
  RemarketingReportAllUnits.Schedule,
  RemarketingReportAllUnits.UnitGroup,
  RemarketingReportAllUnits.Unitnum AS [Cust Unit#],
  RemarketingReportAllUnits.UnitStatus,
  RemarketingReportAllUnits.UnitYr,
  RemarketingReportAllUnits.UnitMake,
  RemarketingReportAllUnits.UnitModel,
  RemarketingReportAllUnits.UnitType,
  RemarketingReportAllUnits.Odometer,
  RemarketingReportAllUnits.UnitVIN,
  RemarketingReportAllUnits.ParkDateProjected,
  RemarketingReportAllUnits.InspectedDate,
  RemarketingReportAllUnits.LocationName,
  RemarketingReportAllUnits.LocationAddress,
  RemarketingReportAllUnits.LCity,
  RemarketingReportAllUnits.LState,
  RemarketingReportAllUnits.TargetSalePrice,
  RemarketingReportAllUnits.UnitSalePrice AS [Actual Sale Price],
  RemarketingReportAllUnits.GLTarget,
  RemarketingReportAllUnits.GLTargetMargin AS VarPct,
  RemarketingReportAllUnits.BuyOutDate,
  RemarketingReportAllUnits.DaysInInventoryPerAccounting AS [Days Since BuyOut],
  RemarketingReportAllUnits.ClientID,
  RemarketingReportAllUnits.BuyerID,
  RemarketingReportAllUnits.UnitID
FROM
  RemarketingReportAllUnits
ORDER BY
  RemarketingReportAllUnits.SourceCompany;
