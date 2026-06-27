SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS RefrigMake,
  ChildParts.ChildClass AS RefrigModel,
  ChildParts.ChildInvoiceNum AS RefrigInvoice,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS Refrig,
  ChildParts.ChildVendorPytDate AS RefrigPaid,
  ChildParts.ChildVIN AS RefrigSerialNum,
  ChildParts.DrawDownId AS RefrigDrawDownId
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Refrig*"
    )
    AND (
      (
        ChildParts.ExcludeCostDocsReports
      )= False
    )
  )
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.ChildVIN,
  ChildParts.DrawDownId;
