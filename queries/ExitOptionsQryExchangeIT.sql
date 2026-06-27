SELECT
  DISTINCTROW UnitExitOption.*,
  Units.UnitUnitNum,
  Units_1.UnitPDStartDate
FROM
  (
    (
      EXITOptions_WorkTable
      INNER JOIN UnitExitOption ON EXITOptions_WorkTable.UnitID = UnitExitOption.UnitId
    )
    INNER JOIN Units ON UnitExitOption.UnitID = Units.UnitID
  )
  LEFT JOIN Units AS Units_1 ON UnitExitOption.ReplacedByUnitID = Units_1.UnitID
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)= 1
    )
  );
