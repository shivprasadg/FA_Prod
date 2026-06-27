SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OilRig,
  ChildParts.ChildMake AS OilRigMake,
  ChildParts.ChildInvoiceNum AS OilRigInvoice,
  ChildParts.ChildVendorPytDate AS OilRigPaid,
  ChildParts.ChildType,
  ChildParts.DrawDownId AS OilRigDraw
FROM
  ChildParts
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.ChildType,
  ChildParts.DrawDownId
HAVING
  (
    (
      (ChildParts.ChildInvoiceNum) Is Not Null
    )
    AND (
      (ChildParts.ChildType) Like "*Oil*"
    )
  );
