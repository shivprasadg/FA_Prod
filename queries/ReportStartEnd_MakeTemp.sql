SELECT
  Units.UnitID,
  Units.UnitStatus,
  Units.ReportStartDate,
  Units.ReportEndDate,
  Units.unitoffleasedt AS OffLease,
  ReMarketing_UnitSales.InvoicePaidDate AS SoldDate,
  ReMarketing_UnitSales.SoldPrice AS SoldPrice,
  ReMarketing_UnitSales.RmktBOSSent AS BOSent INTO ReportStartEnd_SoldUnits
FROM
  Units
  INNER JOIN ReMarketing_UnitSales ON Units.UnitVIN = ReMarketing_UnitSales.unitvin
WHERE
  (
    (
      (Units.ReportEndDate)= #12/31/2030#
    )
  );
