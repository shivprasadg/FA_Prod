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
  Sum(
    ScheduleDocsReportExcelBase.VendorCost
  ) AS VendorCost,
  Sum(
    ScheduleDocsReportExcelBase.SchRent
  ) AS SchRent,
  Sum(
    [ScheduleDocsReportExcelBase].[schrent]
  )/ Sum(
    [ScheduleDocsReportExcelBase].[OEC]
  ) AS SchRentLRF,
  Sum(
    ScheduleDocsReportExcelBase.ExtdRent
  ) AS ExtdRent,
  Sum(
    [ScheduleDocsReportExcelBase].[Extdrent]
  )/ Sum(
    [ScheduleDocsReportExcelBase].[OEC]
  ) AS ExtRentLRF,
  Max(
    ScheduleDocsReportExcelBase.InvoicePaid
  ) AS InvoicePaid,
  ScheduleDocsReportExcelBase.LeaseTermPrimary,
  ScheduleDocsReportExcelBase.LeasePaidAdvArr,
  ScheduleDocsReportExcelBase.[Neg BLCD],
  ScheduleDocsReportExcelBase.[PerDiem Start],
  ScheduleDocsReportExcelBase.BLCD_Projected,
  ScheduleDocsReportExcelBase.Billing,
  ScheduleDocsReportExcelBase.BillingInstructions,
  ScheduleDocsReportExcelBase.SalesTax,
  ScheduleDocsReportExcelBase.RentReceipt,
  ScheduleDocsReportExcelBase.AvalException,
  ScheduleDocsReportExcelBase.TaxInstructions,
  ScheduleDocsReportExcelBase.IRP,
  ScheduleDocsReportExcelBase.FATitleLogNum,
  ScheduleDocsReportExcelBase.SchID,
  Units.UnitEstDelDate AS EstDelv,
  Units.UnitActualDelDate AS ActualDelv,
  Units.UNITPDRENT AS PDRent,
  Units.UnitPDStartDate AS PDStart
FROM
  ScheduleDocsReportExcelBase
  INNER JOIN Units ON ScheduleDocsReportExcelBase.Units = Units.UnitID
GROUP BY
  ScheduleDocsReportExcelBase.Schedule,
  ScheduleDocsReportExcelBase.UnitGroup,
  ScheduleDocsReportExcelBase.Lessee,
  ScheduleDocsReportExcelBase.FAID,
  ScheduleDocsReportExcelBase.Units,
  ScheduleDocsReportExcelBase.UnitNum,
  ScheduleDocsReportExcelBase.VIN,
  ScheduleDocsReportExcelBase.UnitDesc,
  ScheduleDocsReportExcelBase.LeaseTermPrimary,
  ScheduleDocsReportExcelBase.LeasePaidAdvArr,
  ScheduleDocsReportExcelBase.[Neg BLCD],
  ScheduleDocsReportExcelBase.[PerDiem Start],
  ScheduleDocsReportExcelBase.BLCD_Projected,
  ScheduleDocsReportExcelBase.Billing,
  ScheduleDocsReportExcelBase.BillingInstructions,
  ScheduleDocsReportExcelBase.SalesTax,
  ScheduleDocsReportExcelBase.RentReceipt,
  ScheduleDocsReportExcelBase.AvalException,
  ScheduleDocsReportExcelBase.TaxInstructions,
  ScheduleDocsReportExcelBase.IRP,
  ScheduleDocsReportExcelBase.FATitleLogNum,
  ScheduleDocsReportExcelBase.SchID,
  Units.UnitEstDelDate,
  Units.UnitActualDelDate,
  Units.UNITPDRENT,
  Units.UnitPDStartDate;
