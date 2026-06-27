SELECT
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  lookup_groupTypes.grpClass
FROM
  ChildParts
  INNER JOIN lookup_groupTypes ON ChildParts.ChildType = lookup_groupTypes.groupType
WHERE
  (
    (
      (lookup_groupTypes.grpClass)= "other"
    )
  );
