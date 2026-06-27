SELECT
  vw_BankBill.*
FROM
  vw_BankBill
WHERE
  (
    (
      (vw_BankBill.BillID)= [Forms]![Billing]![BillID]
    )
  );
