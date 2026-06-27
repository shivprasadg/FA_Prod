SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS OBC_Make,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OBC,
  ChildParts.ChildInvoiceNum AS OBCInvoice,
  ChildParts.ChildVendorPytDate AS OBCPaid,
  ChildParts.DrawDownId AS OBCDraw,
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
      (ChildParts.ChildType) Like "*OBC*"
      And (ChildParts.ChildType) Not Like "*Install*"
      And (ChildParts.ChildType) Not Like "*Display*"
      And (ChildParts.ChildType) Not Like "*Part*"
    )
  );
