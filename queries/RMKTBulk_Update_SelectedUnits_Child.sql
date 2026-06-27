SELECT
  DISTINCTROW ChildParts.*,
  Units.unitunitnum,
  Units.unitvin,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] AS [Desc],
  Invoices.InvoiceDate,
  [ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass] AS UID
FROM
  (
    (
      Units
      INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
    )
    LEFT JOIN Invoices ON ChildParts.ChildInvRef = Invoices.Invoiceid
  )
  INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID
WHERE
  (
    (
      (ChildParts.ChildVendorREF) Is Not Null
    )
  );
