SELECT
  ChildParts.UnitREF AS AssetId,
  Vendors.VendorSNm AS TaxName,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS FeesTax,
  ChildParts.ChildInvoiceNum AS TaxInvoice,
  ChildParts.ChildVendorPytDate AS TaxPaid,
  ChildParts.DrawDownId AS TaxDraw
FROM
  ChildParts
  INNER JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Tax*"
    )
  )
GROUP BY
  ChildParts.UnitREF,
  Vendors.VendorSNm,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.DrawDownId;
