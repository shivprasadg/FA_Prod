SELECT
  ScheduleDocsReportExcelBase.Schedule,
  ScheduleDocsReportExcelBase.UnitGroup,
  ScheduleDocsReportExcelBase.Lessee,
  ScheduleDocsReportExcelBase.FAID,
  ScheduleDocsReportExcelBase.Units AS AssetID,
  ScheduleDocsReportExcelBase.UnitNum,
  ScheduleDocsReportExcelBase.VIN,
  ScheduleDocsReportExcelBase.UnitDesc,
  Sum(
    ScheduleDocsReportExcelBase.UnitCost
  ) AS UnitCost,
  Sum(
    ScheduleDocsReportExcelBase.[Child Cost]
  ) AS [Child Cost],
  Sum(
    ScheduleDocsReportExcelBase.OEC
  ) AS OEC,
  Max(
    ScheduleDocsReportExcelBase.InvoicePaid
  ) AS InvoicePaid,
  Sum(
    ScheduleDocsReportExcelBase.VendorCost
  ) AS VendorCost,
  ScheduleDocsReportExcelBase.[Allowed Miles],
  ScheduleDocsReportExcelBase.MileChrg,
  ScheduleDocsReportExcelBase.MilesThreshold,
  ScheduleDocsReportExcelBase.[MileChrg After],
  ScheduleDocsReportExcelBase.[Term Year],
  ScheduleDocsReportExcelBase.Rate,
  ScheduleDocsReportExcelBase.Date,
  ScheduleDocsReportExcelBase.[Neg BLCD],
  ScheduleDocsReportExcelBase.[PerDiem Start],
  ScheduleDocsReportExcelBase.BLCD_Projected,
  ScheduleDocsReportExcelBase.Assignee,
  ScheduleDocsReportExcelBase.Spread,
  ScheduleDocsReportExcelBase.Yield,
  ScheduleDocsReportExcelBase.Adjustment,
  ScheduleDocsReportExcelBase.SchID
FROM
  ScheduleDocsReportExcelBase
GROUP BY
  ScheduleDocsReportExcelBase.Schedule,
  ScheduleDocsReportExcelBase.UnitGroup,
  ScheduleDocsReportExcelBase.Lessee,
  ScheduleDocsReportExcelBase.FAID,
  ScheduleDocsReportExcelBase.Units,
  ScheduleDocsReportExcelBase.UnitNum,
  ScheduleDocsReportExcelBase.VIN,
  ScheduleDocsReportExcelBase.UnitDesc,
  ScheduleDocsReportExcelBase.[Allowed Miles],
  ScheduleDocsReportExcelBase.MileChrg,
  ScheduleDocsReportExcelBase.MilesThreshold,
  ScheduleDocsReportExcelBase.[MileChrg After],
  ScheduleDocsReportExcelBase.[Term Year],
  ScheduleDocsReportExcelBase.Rate,
  ScheduleDocsReportExcelBase.Date,
  ScheduleDocsReportExcelBase.[Neg BLCD],
  ScheduleDocsReportExcelBase.[PerDiem Start],
  ScheduleDocsReportExcelBase.BLCD_Projected,
  ScheduleDocsReportExcelBase.Assignee,
  ScheduleDocsReportExcelBase.Spread,
  ScheduleDocsReportExcelBase.Yield,
  ScheduleDocsReportExcelBase.Adjustment,
  ScheduleDocsReportExcelBase.SchID;
