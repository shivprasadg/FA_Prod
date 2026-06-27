SELECT
  FABankBillGroup_MBM.LocationName,
  FABankBillGroup_MBM.SchNo,
  FABankBillGroup_MBM.SGrpNO AS GroupNo,
  FABankBillGroup_MBM.LState,
  FABankBillGroup_MBM.UnitTaxExempt,
  FABankBillGroup_MBM.StTaxType,
  FABankBillGroup_MBM.LTaxRate,
  Sum(FABankBillGroup_MBM.UnitCount) AS UnitCount,
  Sum(FABankBillGroup_MBM.LineTot) AS LineTot,
  Sum(FABankBillGroup_MBM.FSLFee) AS FSLFee,
  Sum(FABankBillGroup_MBM.RentNoTax) AS RentNoTax,
  Sum(FABankBillGroup_MBM.TaxAmt) AS TaxAmt,
  Sum(
    FABankBillGroup_MBM.UnitTotalRent
  ) AS UnitTotalRent,
  Sum(0) AS FSLTax
FROM
  FABankBillGroup_MBM
GROUP BY
  FABankBillGroup_MBM.LocationName,
  FABankBillGroup_MBM.SchNo,
  FABankBillGroup_MBM.SGrpNO,
  FABankBillGroup_MBM.LState,
  FABankBillGroup_MBM.UnitTaxExempt,
  FABankBillGroup_MBM.StTaxType,
  FABankBillGroup_MBM.LTaxRate;
