SELECT
  FARentBillAssignedQ.*
FROM
  FARentBillAssignedQ
WHERE
  (
    (
      (FARentBillAssignedQ.GroupID)= [Forms]![Billing]![GroupId]
    )
  );
