SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS OBC_Display,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OBCDisplay,
  ChildParts.ChildInvoiceNum AS OBCDisplayInv,
  ChildParts.ChildVendorPytDate AS OBCDisplayPaid,
  ChildParts.DrawDownId AS OBCDisplayDraw,
  ChildParts.ChildType
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
      (ChildParts.ChildType) Like "*Display*"
      And (ChildParts.ChildType) Not Like "*Installation*"
    )
  );
