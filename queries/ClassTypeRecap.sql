SELECT
  DISTINCTROW lookup_groupTypes.grpClass,
  lookup_groupTypes.groupType,
  lookup_groupTypes.GrpTypAbrev,
  CLASS.TITLED,
  *
FROM
  CLASS
  INNER JOIN lookup_groupTypes ON CLASS.CLASS = lookup_groupTypes.grpClass
ORDER BY
  lookup_groupTypes.grpClass,
  lookup_groupTypes.groupType;
