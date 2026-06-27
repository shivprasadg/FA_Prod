SELECT
  ScheduleDocsReportExcelBase.Schedule,
  ScheduleDocsReportExcelBase.UnitGroup,
  ScheduleDocsReportExcelBase.Lessee,
  ScheduleDocsReportExcelBase.FAID,
  Count(
    ScheduleDocsReportExcelBase.Units
  ) AS Units,
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
  ScheduleDocsReportExcelBase.LeaseTermPrimary,
  ScheduleDocsReportExcelBase.LeasePaidAdvArr,
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
  ScheduleDocsReportExcelBase.DeliveryLocation,
  ScheduleDocsReportExcelBase.Title,
  ScheduleDocsReportExcelBase.Domicle,
  ScheduleDocsReportExcelBase.Billing,
  ScheduleDocsReportExcelBase.BillingInstructions,
  ScheduleDocsReportExcelBase.SalesTax,
  ScheduleDocsReportExcelBase.RentReceipt,
  ScheduleDocsReportExcelBase.AvalException,
  ScheduleDocsReportExcelBase.TaxInstructions,
  ScheduleDocsReportExcelBase.IRP,
  ScheduleDocsReportExcelBase.SchID
FROM
  ScheduleDocsReportExcelBase
GROUP BY
  ScheduleDocsReportExcelBase.Schedule,
  ScheduleDocsReportExcelBase.UnitGroup,
  ScheduleDocsReportExcelBase.Lessee,
  ScheduleDocsReportExcelBase.FAID,
  ScheduleDocsReportExcelBase.UnitDesc,
  ScheduleDocsReportExcelBase.LeaseTermPrimary,
  ScheduleDocsReportExcelBase.LeasePaidAdvArr,
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
  ScheduleDocsReportExcelBase.DeliveryLocation,
  ScheduleDocsReportExcelBase.Title,
  ScheduleDocsReportExcelBase.Domicle,
  ScheduleDocsReportExcelBase.Billing,
  ScheduleDocsReportExcelBase.BillingInstructions,
  ScheduleDocsReportExcelBase.SalesTax,
  ScheduleDocsReportExcelBase.RentReceipt,
  ScheduleDocsReportExcelBase.AvalException,
  ScheduleDocsReportExcelBase.TaxInstructions,
  ScheduleDocsReportExcelBase.IRP,
  ScheduleDocsReportExcelBase.SchID;
