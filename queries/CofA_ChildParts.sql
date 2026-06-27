SELECT
  ChildParts.UnitREF AS AssetId,
  ChildParts.ChildYear,
  ChildParts.ChildType,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildVIN,
  ChildParts.ChildInvoiceNum,
  Nz([childFinalCost], 0) AS Cost,
  ChildParts.ChildVendorPytDate AS DatePaid,
  ChildParts.REFENG AS EngSerialNum,
  ChildParts.DrawDownId AS RefrigDrawDownId,
  ChildParts.ChildEXCofa,
  ChildParts.ChildPartID
FROM
  ChildParts
WHERE
  (
    (
      (ChildParts.ChildEXCofa)= No
    )
  );
