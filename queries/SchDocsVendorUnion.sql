SELECT
  UnitID,
  UnitVendor as VendorName,
  UnitCost as Cost,
  UnitInvoiceNum as InvoiceNum,
  UnitType,
  VendorID
FROM
  SchDocsVendorUnit
WHERE
  SchID = 3056
UNION ALL
Select
  UnitID,
  ChildVendor,
  ChildCost,
  ChildInvoiceNum,
  ChildType,
  ChildVendorID
FROM
  SchDocsVendorChild
WHERE
  SchID = 3056
ORDER BY
  UnitId,
  Cost DESC,
  UnitType;
