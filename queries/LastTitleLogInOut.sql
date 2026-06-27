SELECT
  UnitTitleLog.UnitRef,
  Max(UnitTitleLog.UnitTitleLogID) AS UnitTitleLogID,
  UnitTitleLog.EntryTypeId
FROM
  UnitTitleLog
GROUP BY
  UnitTitleLog.UnitRef,
  UnitTitleLog.EntryTypeId
ORDER BY
  UnitTitleLog.UnitRef;
