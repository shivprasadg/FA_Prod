SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildMake AS FreightName,
  Sum(
    CCur(
      Nz([childFinalCost], 0)
    )
  ) AS Freight,
  ChildParts.ChildInvoiceNum AS FGTInvoice,
  ChildParts.ChildVendorPytDate AS FreightPaid,
  ChildParts.DrawDownId AS FreightDraw
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
      (ChildParts.ChildType)= "FREIGHT"
    )
  );
