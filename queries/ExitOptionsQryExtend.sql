SELECT
  UnitExitOption.*,
  EXITOptions_WorkTable.UnitUnitNum,
  EXITOptions_WorkTable.GroupID
FROM
  EXITOptions_WorkTable
  INNER JOIN UnitExitOption ON EXITOptions_WorkTable.UnitID = UnitExitOption.UnitID
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)> 2
    )
  )
ORDER BY
  UnitExitOption.UnitID;
