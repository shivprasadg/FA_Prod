SELECT
  DrawDownAvailChildParts.FAID,
  DrawDownAvailChildParts.SchID,
  Vendors.VendorID,
  Vendors.VendorSNm AS Vendor,
  DrawDownAvailChildParts.ChildMake AS Make,
  DrawDownAvailChildParts.ChildType AS Type,
  DrawDownAvailChildParts.GroupID
FROM
  DrawDownAvailChildParts
  INNER JOIN Vendors ON DrawDownAvailChildParts.ChildVendorID = Vendors.VendorID
GROUP BY
  DrawDownAvailChildParts.FAID,
  DrawDownAvailChildParts.SchID,
  Vendors.VendorID,
  Vendors.VendorSNm,
  DrawDownAvailChildParts.ChildMake,
  DrawDownAvailChildParts.ChildType,
  DrawDownAvailChildParts.GroupID
ORDER BY
  DrawDownAvailChildParts.ChildMake;
