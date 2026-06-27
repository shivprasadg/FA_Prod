SELECT
  ChildParts.UnitREF AS AssetId,
  Last(ChildParts.ChildMake) AS WarrantyName,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS Warranty,
  ChildParts.ChildInvoiceNum AS WarrantyInv,
  ChildParts.ChildVendorPytDate AS WarrantyPaid,
  ChildParts.DrawDownId AS WarrantyDraw
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Warranty*"
    )
  )
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.DrawDownId;
