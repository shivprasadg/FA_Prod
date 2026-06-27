SELECT
  Count(ChildParts.ChildPartID) AS CountOfChildPartID,
  IIf(
    Nz([ChildType], "")= "",
    Null,
    [ChildType] & " | "
  )& IIf(
    Nz([ChildYear], 0)= 0,
    Null,
    [ChildYear] & " "
  )& IIf(
    Nz([ChildMake], "")= "",
    Null,
    [ChildMake] & " "
  )& IIf(
    Nz([ChildClass], "")= "",
    Null,
    [ChildClass] & " "
  ) AS PartDescription,
  ChildParts.childFinalCost,
  ChildParts.PONumber,
  [childFinalCost] * Count([ChildPartID]) AS ExtendedCost
FROM
  ChildParts
  INNER JOIN vw_SixKeys ON (
    ChildParts.UnitREF = vw_SixKeys.UnitID
  )
  AND (
    ChildParts.PONumber = vw_SixKeys.PurchaseOrder
  )
GROUP BY
  IIf(
    Nz([ChildType], "")= "",
    Null,
    [ChildType] & " | "
  )& IIf(
    Nz([ChildYear], 0)= 0,
    Null,
    [ChildYear] & " "
  )& IIf(
    Nz([ChildMake], "")= "",
    Null,
    [ChildMake] & " "
  )& IIf(
    Nz([ChildClass], "")= "",
    Null,
    [ChildClass] & " "
  ),
  ChildParts.childFinalCost,
  ChildParts.PONumber;
