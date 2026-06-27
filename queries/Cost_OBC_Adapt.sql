SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS OBC_Adapt,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OBCAdapt,
  ChildParts.ChildInvoiceNum AS OBCAdaptInv,
  ChildParts.ChildVendorPytDate AS OBCAdaptPaid,
  ChildParts.DrawDownId AS OBCAdaptDraw,
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
      (ChildParts.ChildType) Like "*Adapt*"
      And (ChildParts.ChildType) Not Like "*Installation*"
    )
  );
