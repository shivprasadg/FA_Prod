SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS OBC_Multi,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OBCMulti,
  ChildParts.ChildInvoiceNum AS OBCMultiInv,
  ChildParts.ChildVendorPytDate AS OBCMultiPaid,
  ChildParts.DrawDownId AS OBCMultiDraw
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
      (ChildParts.ChildType) Like "*Multi*"
      And (ChildParts.ChildType) Not Like "*Installation*"
    )
  );
