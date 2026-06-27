CREATE TABLE [ReportStartEnd_SoldUnits] (
  [UnitID] AUTOINCREMENT,
  [UnitStatus] VARCHAR (2),
  [ReportStartDate] DATETIME,
  [ReportEndDate] DATETIME,
  [OffLease] DATETIME,
  [SoldDate] DATETIME,
  [SoldPrice] CURRENCY,
  [BOSent] DATETIME
)
