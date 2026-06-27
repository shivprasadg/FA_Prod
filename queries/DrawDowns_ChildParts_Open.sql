SELECT
  ChildParts.ChildPartID,
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.Unitnum,
  ChildParts.ChildInvoiceNum AS CInvoice,
  ChildParts.ChildType,
  ChildParts.childFinalCost AS PCost,
  ChildParts.DrawDownId,
  Vendors.vendorsnm AS CVendor,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  DrawDowns.DrawRequestNumber AS RequestNumber,
  DrawDowns.DrawName,
  [ChildMake] & " " & [ChildClass] & " " & [ChildType] AS ChildDesc,
  Vendors.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID AS GroupID
FROM
  (
    Vendors
    INNER JOIN (
      vw_SixKeys
      INNER JOIN ChildParts ON vw_SixKeys.UnitID = ChildParts.UnitREF
    ) ON Vendors.VendorID = ChildParts.ChildVendorREF
  )
  LEFT JOIN DrawDowns ON vw_SixKeys.DrawDownId = DrawDowns.DrawDownId
WHERE
  (
    (
      (ChildParts.DrawDownId) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  Vendors.vendorsnm;
