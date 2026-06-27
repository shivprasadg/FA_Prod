SELECT
  Lookup_BillType.BillTypeID,
  Lookup_BillType.BillType,
  Lookup_BillType.Obsolete
FROM
  Lookup_BillType
WHERE
  (
    (
      (Lookup_BillType.Obsolete)= False
    )
  )
ORDER BY
  Lookup_BillType.BillType;
