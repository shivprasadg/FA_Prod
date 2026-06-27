SELECT
  vw_RentInvoice.BillID,
  vw_RentInvoice.GroupID,
  vw_RentInvoice.Schedule,
  vw_RentInvoice.UnitGroup,
  vw_RentInvoice.LocationName,
  vw_RentInvoice.LCity,
  vw_RentInvoice.LState,
  vw_RentInvoice.StTaxType,
  Sum(vw_RentInvoice.CountMe) AS UnitCount,
  Sum(vw_RentInvoice.BilledRent) AS BilledRent,
  Sum(vw_RentInvoice.BilledFees) AS BilledFees,
  Sum(vw_RentInvoice.UnitTaxes) AS UnitTaxes,
  Sum(vw_RentInvoice.UnitTotalRent) AS UnitTotalRent
FROM
  vw_RentInvoice
GROUP BY
  vw_RentInvoice.BillID,
  vw_RentInvoice.GroupID,
  vw_RentInvoice.Schedule,
  vw_RentInvoice.UnitGroup,
  vw_RentInvoice.LocationName,
  vw_RentInvoice.LCity,
  vw_RentInvoice.LState,
  vw_RentInvoice.StTaxType;
