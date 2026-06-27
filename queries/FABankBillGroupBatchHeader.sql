SELECT
  FABankBillGroupBatch.BillID,
  FABankBillGroupBatch.BillNo,
  FABankBillGroupBatch.BankName,
  FABankBillGroupBatch.BillGroupID,
  FABankBillGroupBatch.BankContract,
  FABankBillGroupBatch.Billfrom,
  FABankBillGroupBatch.Billto,
  FABankBillGroupBatch.SchNo,
  FABankBillGroupBatch.SGrpNO AS GroupNo,
  Sum(FABankBillGroupBatch.UnitCount) AS UnitCount,
  Sum(FABankBillGroupBatch.LineTot) AS LineTot,
  Sum(FABankBillGroupBatch.FSLFee) AS FSLFee,
  Sum(FABankBillGroupBatch.RentNoTax) AS RentNoTax,
  Sum(FABankBillGroupBatch.TaxAmt) AS TaxAmt,
  Sum(
    FABankBillGroupBatch.UnitTotalRent
  ) AS UnitTotalRent,
  Sum(0) AS FSLTax,
  Sum(FABankBillGroupBatch.CurrRent) AS SumOfCurrRent,
  FABankBillGroupBatch.BillToSchedule,
  FABankBillGroupBatch.RemitAssignee,
  FABankBillGroupBatch.aSort
FROM
  FABankBillGroupBatch
GROUP BY
  FABankBillGroupBatch.BillID,
  FABankBillGroupBatch.BillNo,
  FABankBillGroupBatch.BankName,
  FABankBillGroupBatch.BillGroupID,
  FABankBillGroupBatch.BankContract,
  FABankBillGroupBatch.Billfrom,
  FABankBillGroupBatch.Billto,
  FABankBillGroupBatch.SchNo,
  FABankBillGroupBatch.SGrpNO,
  FABankBillGroupBatch.BillToSchedule,
  FABankBillGroupBatch.RemitAssignee,
  FABankBillGroupBatch.aSort;
