SELECT
  FABankBill.*,
  FABankBill.UNITSRENT AS RentNoTax,
  IIf(
    [UnitTaxExempt] = False,
    CCur([FSLFee] * [LTaxRate]),
    0
  ) AS FSLTax,
  [RentNoTax] + CCur(
    Nz([TaxAmt], 0)
  )+ [FLSFEE] AS UnitTotalRent,
  [BLocStreet] & Chr(13)& Chr(10)& [BLocSt2] & Chr(13)& Chr(10)& IIf(
    IsNull([BLoST3]),
    Null,
    [BLoST3] & Chr(13)& Chr(10)
  )& [blOCcITY] & ", " & [bLOCST] & " " & Left([BlOCzIP], 5)& "-" & Right([BlOCzIP], 4) AS BillToAddress,
  FABankBill.BillNo AS InvoiceNumber
FROM
  FABankBill
WHERE
  (
    (
      (FABankBill.SGrpID)= Forms!Billing!GroupId
    )
    And (
      (FABankBill.LocationID)<> 1173
    )
  );
