SELECT
  FABankBillReg.LocationName,
  FABankBillReg.SchNo,
  FABankBillReg.SGrpNO AS GroupNo,
  FABankBillReg.LState,
  FABankBillReg.UnitTaxExempt,
  FABankBillReg.StTaxType,
  FABankBillReg.LTaxRate,
  Count(FABankBillReg.AssetID) AS UnitCount,
  Sum(FABankBillReg.LineTot) AS LineTot,
  Sum(FABankBillReg.FSLFee) AS FSLFee,
  Sum(FABankBillReg.RentNoTax) AS RentNoTax,
  Sum(FABankBillReg.TaxAmt) AS TaxAmt,
  Sum(FABankBillReg.UnitTotalRent) AS UnitTotalRent,
  Sum(0) AS FSLTax,
  FABankBillReg.BillMonths
FROM
  FABankBillReg
GROUP BY
  FABankBillReg.LocationName,
  FABankBillReg.SchNo,
  FABankBillReg.SGrpNO,
  FABankBillReg.LState,
  FABankBillReg.UnitTaxExempt,
  FABankBillReg.StTaxType,
  FABankBillReg.LTaxRate,
  FABankBillReg.BillMonths;
