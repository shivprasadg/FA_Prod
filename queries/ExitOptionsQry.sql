SELECT
  DISTINCTROW UnitExitOption.*,
  Units.UnitUnitNum
FROM
  (
    EXITOptions_WorkTable
    INNER JOIN UnitExitOption ON EXITOptions_WorkTable.UnitID = UnitExitOption.UnitId
  )
  INNER JOIN Units ON UnitExitOption.UnitID = Units.UnitID
WHERE
  (
    (
      (UnitExitOption.OptionTypeDesc)= "Extend"
    )
  );
