SELECT
  ChildParts.ChildPartID,
  ChildParts.ChildYear,
  ChildParts.ChildMakeRef,
  ChildParts.ChildMake,
  ChildParts.ChildType,
  ChildParts.ChildStatus,
  ChildParts.ChildVIN,
  ChildParts.childFinalCost,
  Units.UnitID,
  ChildParts.UnitREF
FROM
  ChildParts
  LEFT JOIN Units ON ChildParts.UnitREF = Units.UnitID
WHERE
  (
    (
      (Units.UnitID) Is Null
    )
  );
