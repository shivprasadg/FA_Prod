SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS OBC_Part,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OBCPart,
  ChildParts.ChildInvoiceNum AS OBCPartInv,
  ChildParts.ChildVendorPytDate AS OBCPartPaid,
  ChildParts.DrawDownId AS OBCPartDraw,
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
      (ChildParts.ChildType) Like "*OBC Part*"
    )
  );
