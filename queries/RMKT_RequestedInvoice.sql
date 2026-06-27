SELECT
  RmktBills.[Invoice No],
  RmktBills.ReqDt,
  RmktBills.BuyerRef
FROM
  RmktBills
WHERE
  (
    (
      (RmktBills.ReqDt)= Date()
    )
  );
