TRANSFORM First(Vendors.VendorSNm) AS ChildVendor
SELECT
  Cost_ChildParts.AssetID
FROM
  Cost_ChildParts
  INNER JOIN Vendors ON Cost_ChildParts.ChildVendor = Vendors.VendorID
WHERE
  (
    (
      (Cost_ChildParts.SchID)= 3056
    )
  )
GROUP BY
  Cost_ChildParts.AssetID PIVOT [PartType] & ' Vendor';
