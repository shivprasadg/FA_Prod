SELECT
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  Units.unitmake,
  lookup_groupTypes.grpClass,
  Count(Units.UnitID) AS CountOfUnitID,
  Units.UnitSubType
FROM
  (
    Schedule
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    Units
    INNER JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
  ) ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  Units.unitmake,
  lookup_groupTypes.grpClass,
  Units.UnitSubType;
