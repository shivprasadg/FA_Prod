SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS OBC_Dome,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OBCDome,
  ChildParts.ChildInvoiceNum AS OBCDomeInv,
  ChildParts.ChildVendorPytDate AS OBCDomePaid,
  ChildParts.DrawDownId AS OBCDomeDraw
FROM
  ChildParts
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.DrawDownId,
  ChildParts.ChildType
HAVING
  (
    (
      (ChildParts.ChildType) Like "*Dome*"
    )
  );
