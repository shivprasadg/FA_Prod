SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS PTO,
  ChildParts.ChildMake AS PTOMake,
  ChildParts.ChildInvoiceNum AS PTOInvoice,
  ChildParts.ChildType,
  ChildParts.ChildVendorPytDate AS PTOPaid,
  ChildParts.DrawDownId AS PTODraw
FROM
  ChildParts
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildType,
  ChildParts.ChildVendorPytDate,
  ChildParts.DrawDownId
HAVING
  (
    (
      (ChildParts.ChildInvoiceNum) Is Not Null
    )
    AND (
      (ChildParts.ChildType) Like "*PTO*"
    )
  );
