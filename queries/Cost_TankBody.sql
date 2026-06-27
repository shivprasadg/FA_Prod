SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS TankBody,
  ChildParts.ChildMake AS TankMake,
  ChildParts.ChildInvoiceNum AS TankInvoice,
  ChildParts.ChildVendorPytDate AS TankPaid,
  ChildParts.ChildType,
  ChildParts.DrawDownId AS TankDraw
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
      (ChildParts.ChildType)= "Tank Body"
    )
  );
