SELECT
  Units.UnitID,
  ChildParts.ChildYear,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.unitref,
  ChildParts.childfinalcost,
  ChildParts.childvendorref
FROM
  Units
  LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF;
