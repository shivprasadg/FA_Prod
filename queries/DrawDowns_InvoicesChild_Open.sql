SELECT
  ChildParts.DrawDownId,
  Vendors.vendorsnm AS CVendor,
  ChildParts.ChildInvoiceNum AS CInvoice,
  Count(ChildParts.ChildPartID) AS ChildParts,
  Sum(ChildParts.childFinalCost) AS CCost,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Vendors.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
FROM
  Vendors
  INNER JOIN (
    vw_SixKeys
    INNER JOIN ChildParts ON vw_SixKeys.UnitID = ChildParts.UnitREF
  ) ON Vendors.VendorID = ChildParts.ChildVendorREF
GROUP BY
  ChildParts.DrawDownId,
  Vendors.vendorsnm,
  ChildParts.ChildInvoiceNum,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Vendors.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
HAVING
  (
    (
      (ChildParts.DrawDownId) Is Null
    )
  )
ORDER BY
  Vendors.vendorsnm;
