SELECT
  UnitTitleLog.UnitRef,
  Max(UnitTitleLog.UnitTitleLogDate) AS MaxOfUnitTitleLogDate,
  Max(UnitTitleLog.UnitTitleLogId) AS MaxOfUnitTitleLogId
FROM
  Units
  INNER JOIN UnitTitleLog ON Units.UnitID = UnitTitleLog.UnitRef
GROUP BY
  UnitTitleLog.UnitRef
ORDER BY
  UnitTitleLog.UnitRef;
