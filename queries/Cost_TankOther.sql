SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS TankOther,
  ChildParts.ChildMake AS TankOthMake,
  ChildParts.ChildInvoiceNum AS TankOthInvoice,
  ChildParts.ChildVendorPytDate AS TankOthPaid,
  ChildParts.ChildType,
  ChildParts.DrawDownId AS TankOthDraw
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
      (ChildParts.ChildType)<> "Tank Body"
      And (ChildParts.ChildType) Like "*Tank*"
    )
  );
