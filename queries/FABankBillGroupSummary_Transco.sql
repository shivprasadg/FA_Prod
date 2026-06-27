SELECT
  FABankBillGroup_Transco.LocationName,
  FABankBillGroup_Transco.SchNo,
  FABankBillGroup_Transco.SGrpNO AS GroupNo,
  FABankBillGroup_Transco.LState,
  FABankBillGroup_Transco.UnitTaxExempt,
  FABankBillGroup_Transco.StTaxType,
  FABankBillGroup_Transco.LTaxRate,
  Sum(
    FABankBillGroup_Transco.UnitCount
  ) AS UnitCount,
  Sum(
    FABankBillGroup_Transco.LineTot
  ) AS LineTot,
  Sum(FABankBillGroup_Transco.FSLFee) AS FSLFee,
  Sum(
    FABankBillGroup_Transco.RentNoTax
  ) AS RentNoTax,
  Sum(FABankBillGroup_Transco.TaxAmt) AS TaxAmt,
  Sum(
    FABankBillGroup_Transco.UnitTotalRent
  ) AS UnitTotalRent,
  Sum(0) AS FSLTax
FROM
  FABankBillGroup_Transco
GROUP BY
  FABankBillGroup_Transco.LocationName,
  FABankBillGroup_Transco.SchNo,
  FABankBillGroup_Transco.SGrpNO,
  FABankBillGroup_Transco.LState,
  FABankBillGroup_Transco.UnitTaxExempt,
  FABankBillGroup_Transco.StTaxType,
  FABankBillGroup_Transco.LTaxRate;
