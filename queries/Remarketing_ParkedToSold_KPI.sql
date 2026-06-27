SELECT
  RemarketingReportAllUnits.SourceCompany,
  RemarketingReportAllUnits.UnitID,
  RemarketingReportAllUnits.[Sch|Grp],
  RemarketingReportAllUnits.Unitnum,
  RemarketingReportAllUnits.UnitYr,
  RemarketingReportAllUnits.UnitMake,
  RemarketingReportAllUnits.UnitModel,
  RemarketingReportAllUnits.UnitType,
  RemarketingReportAllUnits.UnitSubType,
  RemarketingReportAllUnits.UnitStatus,
  RemarketingReportAllUnits.InspectedDate,
  RemarketingReportAllUnits.[SUF Sent],
  RemarketingReportAllUnits.[SUF Returned],
  RemarketingReportAllUnits.ParkDate,
  RemarketingReportAllUnits.OffLeaseDate,
  RemarketingReportAllUnits.BuyOutDate,
  RemarketingReportAllUnits.InvoicePaid,
  IIf(
    IsDate([ParkDate])
    And IsDate([BuyOutdate])
    And ([ParkDate] < [BuyOutdate]),
    [BuyOutdate] - [ParkDate],
    Null
  ) AS [Parked To BuyOut],
  IIf(
    IsDate([InvoicePaid])
    And IsDate([BuyOutdate])
    And [InvoicePaid] >= [BuyOutDate],
    [InvoicePaid] - [Buyoutdate],
    IIf(
      [InvoicePaid] < [BuyOutDate], 0, Null
    )
  ) AS [BuyOut to Sold],
  IIf(
    IsDate([ParkDate])
    And IsDate([InvoicePaid]),
    [InvoicePaid] - [Parkdate],
    Null
  ) AS [Parked To Sold],
  IIf(
    Not IsDate([InvoicePaid])
    And IsDate([BuyOutdate]),
    Date()- [BuyOutdate]
  ) AS [Bought Still in Inv],
  IIf(
    IsDate([ParkDate])
    And IsDate([OffLeaseDate]),
    [ParkDate] - [OffLeaseDate],
    Null
  ) AS [Parked to OffLease],
  IIf(
    IsDate([InvoicePaid])
    And IsDate([OffLeaseDate]),
    [InvoicePaid] - [OffLeaseDate],
    Null
  ) AS [OffLease to Sold]
FROM
  RemarketingReportAllUnits;
