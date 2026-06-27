SELECT
  BillAdjT.BillAdjID,
  BillAdjT.BillAdjBillref,
  BillAdjT.BillAdj,
  BillAdjT.BillAdjtype,
  BillAdjT.BillAdjVIN,
  BillAdjT.BillAdjUnitNo,
  BillAdjT.BillAdjNote
FROM
  BillAdjT
WHERE
  (
    (
      (BillAdjT.BillAdjBillref)= [forms]![BillingSubForm]![billid]
    )
  );
