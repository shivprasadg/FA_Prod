SELECT
  FABankBill.*,
  FABankBill.CurrRent AS RentNoTax,
  0 AS FSLTax,
  [RentNoTax] + CCur(
    Nz([FSLFee], 0)
  )+ CCur(
    Nz([TaxAmt], 0)
  ) AS UnitTotalRent,
  FABankBill.BillToSchedule AS BillToAddress,
  FABankBill.BillNo AS InvoiceNumber
FROM
  vw_BankBill AS FABankBill
WHERE
  (
    (
      (FABankBill.BillID)= Forms!Billing!BillID
    )
    And (
      (FABankBill.LocationID) Is Not Null
    )
    And (
      (FABankBill.SGrpID) Is Not Null
    )
  );
