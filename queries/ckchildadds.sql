SELECT
  ChildParts.UnitREF,
  ChildParts.ChildPartID,
  ChildParts.ChildMake,
  ChildParts.childFinalCost
FROM
  ChildParts
  INNER JOIN Units ON ChildParts.UnitREF = Units.UnitID
WHERE
  (
    (
      (ChildParts.ChildMake)= "vomela"
    )
    AND (
      (ChildParts.childFinalCost)= 351.37
    )
    AND (
      (Units.GroupREF)= 1417
    )
  )
ORDER BY
  ChildParts.UnitREF,
  ChildParts.ChildPartID DESC;
