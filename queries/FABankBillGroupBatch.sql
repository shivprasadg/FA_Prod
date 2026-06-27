SELECT
  vw_BankBill_Batch.*,
  CurrRent AS RentNoTax,
  IIf(
    [UnitTaxExempt] = False,
    CCur([FSLFee] * [LTaxRate]),
    0
  ) AS FSLTax,
  [CurrRent] + CCur(
    Nz([FSLFee], 0)
  )+ CCur(
    Nz([TaxAmt], 0)
  ) AS UnitTotalRent,
  BillToGroup AS BillToAddress,
  BillNo AS InvoiceNumber
FROM
  vw_BankBill_Batch
WHERE
  BillID In (36718);
