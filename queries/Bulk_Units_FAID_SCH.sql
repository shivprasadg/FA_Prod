SELECT
  DISTINCTROW Schedule.SchFAID,
  SchGrp.SGrpNO,
  Schedule.SchNo,
  Units.UnitStatus,
  Units.unitID AS AssetID,
  [UnitYr] & ' ' & [UnitMake] & ' ' & [UnitSubType] & ' ' & [UnitType] & IIf(
    Len([UnitConfiguration])> 0,
    " " & [UnitConfiguration],
    ""
  )& IIf(
    Len([UnitDescriptor])> 0,
    " " & [UnitDescriptor],
    ""
  )& " (" & [UnitStatus] & ")" AS UnitDesc,
  Units.UnitYr,
  Units.unitmake,
  Units.unittype,
  Units.UnitVIN,
  Units.UnitUnitNum AS UnitNum,
  Units.UnitVINLocked,
  SchGrp.SGrpID,
  Schedule.SchID,
  Units.UnitConfiguration,
  Units.UnitDescriptor,
  Units.DrawDownID
FROM
  (
    Schedule
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF;
