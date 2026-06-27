SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS Connect_Make,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS [Connect],
  ChildParts.ChildInvoiceNum AS ConnectInvoice,
  ChildParts.ChildVendorPytDate AS ConnectPaid,
  ChildParts.DrawDownId AS ConnectDraw
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
      (ChildParts.ChildType) Like "*Connection*"
    )
  );
