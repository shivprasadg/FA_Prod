SELECT
  UnitExitOption.*,
  Units.UnitUnitNum,
  Units.unitoffleasedt,
  Units.UnitOLInd,
  Units.UnitStatus
FROM
  (
    EXITOptions_WorkTable
    INNER JOIN UnitExitOption ON EXITOptions_WorkTable.UnitID = UnitExitOption.UnitId
  )
  INNER JOIN Units ON UnitExitOption.UnitID = Units.UnitID
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)= 2
    )
  );
