SELECT
  UnitExitOption.UnitID
FROM
  EXITOptions_WorkTable
  LEFT JOIN UnitExitOption ON (
    EXITOptions_WorkTable.UID0 = UnitExitOption.UIDX
  )
  AND (
    EXITOptions_WorkTable.UnitID = UnitExitOption.UnitID
  )
WHERE
  (
    (
      (UnitExitOption.UnitID) Is Null
    )
  );
