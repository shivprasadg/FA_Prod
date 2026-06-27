SELECT
  FABankBill.*,
  FABankBill.CurrRent AS RentNoTax,
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
  FABankBill.BillToGroup AS BillToAddress,
  FABankBill.BillNo AS InvoiceNumber,
  FABankBill.SGrpID
FROM
  FABankBill
WHERE
  (
    (
      (FABankBill.SGrpID)= [Forms]![Billing]![GroupId]
    )
  );
