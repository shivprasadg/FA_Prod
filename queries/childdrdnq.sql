SELECT
  Units.UnitID,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.ChildInvoiceNum,
  ChildParts.childFinalCost,
  Vendors.VendorName,
  ChildParts.ChildVendorPytDate
FROM
  Vendors
  INNER JOIN (
    (
      (
        Schedule
        INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
      )
      INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
    )
    INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
  ) ON Vendors.VendorID = ChildParts.ChildVendorREF
WHERE
  (
    (
      (ChildParts.childFinalCost)> 0
    )
    AND (
      (ChildParts.ChildVendorPytDate) Is Null
    )
    AND (
      (ChildParts.ExclFrGrpOEC)= False
    )
  )
ORDER BY
  Units.UnitID,
  ChildParts.ChildType;
