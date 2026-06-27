SELECT
  DISTINCTROW Schedule.SchFAID,
  SchGrp.SGrpNO,
  Schedule.SchNo,
  Units.UnitStatus,
  Units.unitID AS AssetID,
  Units.unitmake,
  Units.unittype
FROM
  (
    Schedule
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Units.UnitStatus)= "A"
    )
  );
