SELECT
  RemarketUnitStatusAudit.*
FROM
  RemarketUnitStatusAudit
WHERE
  (
    (
      (
        RemarketUnitStatusAudit.BuyerId
      ) Is Not Null
    )
    AND (
      (RemarketUnitStatusAudit.Status) Not Like "*S*"
      And (RemarketUnitStatusAudit.Status) Not Like "*R*"
      And (RemarketUnitStatusAudit.Status)<> "D"
    )
    AND (
      (
        RemarketUnitStatusAudit.unitoffleasedt
      ) Is Null
    )
  )
  OR (
    (
      (
        RemarketUnitStatusAudit.InvoicePaidDate
      ) Is Null
    )
    AND (
      (
        RemarketUnitStatusAudit.BuyerId
      ) Is Not Null
    )
    AND (
      (RemarketUnitStatusAudit.Status)= "S"
      Or (RemarketUnitStatusAudit.Status)= "KS"
      Or (RemarketUnitStatusAudit.Status)= "D"
    )
    AND (
      (
        RemarketUnitStatusAudit.unitoffleasedt
      ) Is Null
    )
  )
  OR (
    (
      (RemarketUnitStatusAudit.Status) Like "*S*"
      Or (RemarketUnitStatusAudit.Status)= "D"
    )
    AND (
      (
        RemarketUnitStatusAudit.unitoffleasedt
      ) Is Null
    )
  );
