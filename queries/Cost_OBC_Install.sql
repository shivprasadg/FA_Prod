SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS OBCInstall,
  ChildParts.ChildInvoiceNum AS OBCInsInvoice,
  Nz([ChildMake], "Other") AS OBCInstaller,
  ChildParts.ChildVendorPytDate AS OBCInstallPaid,
  ChildParts.DrawDownId AS OBCInstallDraw,
  ChildParts.ChildType
FROM
  ChildParts
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildInvoiceNum,
  Nz([ChildMake], "Other"),
  ChildParts.ChildVendorPytDate,
  ChildParts.DrawDownId,
  ChildParts.ChildType
HAVING
  (
    (
      (ChildParts.ChildType) Like "*OBC*"
      And (ChildParts.ChildType) Like "*Install*"
    )
  );
