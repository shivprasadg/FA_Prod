TRANSFORM First(
  Nz([PartInvoice], "-")
) AS Expr1
SELECT
  Cost_ChildParts.AssetID
FROM
  Cost_ChildParts
  INNER JOIN Vendors ON Cost_ChildParts.ChildVendor = Vendors.VendorID
WHERE
  (
    (
      (Cost_ChildParts.SchID)= 1277
    )
  )
GROUP BY
  Cost_ChildParts.AssetID PIVOT 'Vendor :' & [VendorSNm];
