SELECT
  FABankBill.*,
  FABankBill.UNITRENT AS RentNoTax,
  IIf(
    [UnitTaxExempt] = False,
    CCur([FSLFee] * [LTaxRate]),
    0
  ) AS FSLTax,
  [RentNoTax] + CCur(
    Nz([FSLFee], 0)
  )+ CCur(
    Nz([TaxAmt], 0)
  ) AS UnitTotalRent,
  "Meadowbrook Meat Company Inc." & Chr(13)& Chr(10)& "PO BOX 6258" & Chr(13)& Chr(10)& "Rocky Mount NC 27802" & Chr(13)& Chr(10)& "Attn: Tiffany Hudson" AS BillToMBMAdr,
  FABankBill.BillNo AS InvoiceNumber,
  FABankBill.BillToGroup AS BillToAddress
FROM
  FABankBill
WHERE
  (
    (
      (FABankBill.LocationID)= 1173
    )
    And (
      (FABankBill.SGrpID)= Forms!Billing!GroupId
    )
  );
