SELECT
  DISTINCTROW ChildParts.*,
  Units.UnitUnitNum,
  Units.UnitVIN,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& IIf(
    [UnitType] = "Tractor", "", " " & [UnitType]
  ) AS [Desc],
  Invoices.InvoiceDate,
  [ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass] AS UID
FROM
  (
    Bulk_Update_WorkTable
    INNER JOIN (
      Units
      INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
    ) ON Bulk_Update_WorkTable.UnitID = ChildParts.UnitREF
  )
  LEFT JOIN Invoices ON ChildParts.ChildInvoiceNum = Invoices.InvoiceNumber
WHERE
  (
    (
      (ChildParts.ChildVendorREF) Is Not Null
    )
  );
