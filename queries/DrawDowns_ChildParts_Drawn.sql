SELECT
  ChildParts.ChildPartID,
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.UnitNum,
  ChildParts.ChildInvoiceNum AS CInvoice,
  ChildParts.ChildType,
  ChildParts.childFinalCost AS CCost,
  ChildParts.DrawDownId,
  Banks.BankShortName AS cIFBank,
  ChildParts.cIFAmount,
  Vendors.vendorsnm AS CVendor,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLNo,
  DrawDowns.DrawRequestNumber AS RequestNumber,
  [ChildMake] & " " & [ChildClass] & " " & [ChildType] AS ChildDesc,
  DrawDowns.DrawName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Vendors.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  DrawDowns.DrawRequestNumber
FROM
  (
    (
      Vendors
      INNER JOIN (
        vw_SixKeys
        INNER JOIN ChildParts ON vw_SixKeys.UnitID = ChildParts.UnitREF
      ) ON Vendors.VendorID = ChildParts.ChildVendorREF
    )
    LEFT JOIN DrawDowns ON vw_SixKeys.DrawDownId = DrawDowns.DrawDownId
  )
  LEFT JOIN Banks ON ChildParts.cIFBankID = Banks.BankID
WHERE
  (
    (
      (ChildParts.DrawDownId) Is Not Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  Vendors.vendorsnm;
