SELECT
  FABankBillGroupBatch.LocationName,
  FABankBillGroupBatch.LState,
  Sum(FABankBillGroupBatch.UnitCount) AS UnitCount,
  Sum(FABankBillGroupBatch.LineTot) AS LineTot,
  Sum(FABankBillGroupBatch.FSLFee) AS FSLFee,
  Sum(FABankBillGroupBatch.RentNoTax) AS RentNoTax,
  Sum(FABankBillGroupBatch.TaxAmt) AS TaxAmt,
  Sum(
    FABankBillGroupBatch.UnitTotalRent
  ) AS UnitTotalRent,
  0 AS FSLTax
FROM
  FABankBillGroupBatch
GROUP BY
  FABankBillGroupBatch.LocationName,
  FABankBillGroupBatch.LState;
