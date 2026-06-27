SELECT
  FABankBillGroup.LocationName,
  FABankBillGroup.SchNo,
  FABankBillGroup.SGrpNO AS GroupNo,
  FABankBillGroup.LState,
  FABankBillGroup.UnitTaxExempt,
  FABankBillGroup.StTaxType,
  FABankBillGroup.LTaxRate,
  Sum(FABankBillGroup.UnitCount) AS UnitCount,
  Sum(FABankBillGroup.LineTot) AS LineTot,
  Sum(FABankBillGroup.FSLFee) AS FSLFee,
  Sum(FABankBillGroup.RentNoTax) AS RentNoTax,
  Sum(FABankBillGroup.TaxAmt) AS TaxAmt,
  Sum(FABankBillGroup.UnitTotalRent) AS UnitTotalRent,
  Sum(0) AS FSLTax,
  FABankBillGroup.CurrRent
FROM
  FABankBillGroup
GROUP BY
  FABankBillGroup.LocationName,
  FABankBillGroup.SchNo,
  FABankBillGroup.SGrpNO,
  FABankBillGroup.LState,
  FABankBillGroup.UnitTaxExempt,
  FABankBillGroup.StTaxType,
  FABankBillGroup.LTaxRate,
  FABankBillGroup.CurrRent;
