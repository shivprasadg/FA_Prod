SELECT
  ChildParts.UnitREF,
  ChildParts.ChildVIN,
  ChildParts.ChildYear,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.EstimatedDeliveryDate,
  lookup_groupTypes.grpClass,
  ChildParts.ActualDeliveryDate
FROM
  ChildParts
  INNER JOIN lookup_groupTypes ON ChildParts.ChildType = lookup_groupTypes.groupType
WHERE
  (
    (
      (lookup_groupTypes.grpClass)= "body"
    )
  )
ORDER BY
  ChildParts.UnitREF;
