SELECT
  ChildParts.UnitREF AS AssetId,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS Inspection,
  ChildParts.ChildMake AS InspCo,
  ChildParts.ChildInvoiceNum AS InspInvoice,
  ChildParts.ChildVendorPytDate AS InspPaid,
  ChildParts.DrawDownId AS InspDraw
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Inspection*"
    )
  )
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.DrawDownId;
