SELECT
  ChildParts.UnitREF,
  ChildParts.ChildVIN AS [Serial Number],
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.ChildInvoiceNum AS RefInvoiceNum
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType)= "refrig unit"
    )
  )
ORDER BY
  ChildParts.UnitREF;
