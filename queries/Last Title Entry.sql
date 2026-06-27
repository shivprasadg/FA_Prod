SELECT
  UnitTitleLog.UnitRef AS AssetID,
  UnitTitleLog.UnitRef,
  Last(UnitTitleLog.UnitTitleLogId) AS TitleLogId,
  Max(UnitTitleLog.UnitTitleLogDate) AS LogDate,
  Last(
    IIf([entrytypeid] = 1, "In", "Out")
  ) AS [Action],
  Last(UnitTitleLog.TitleNumber) AS LastOfTitleNumber,
  Max(UnitTitleLog.UnitTitleLogDate) AS LastOfUnitTitleLogDate,
  LookUp_TitleReasons.ReasonDescription
FROM
  UnitTitleLog
  LEFT JOIN LookUp_TitleReasons ON UnitTitleLog.EntryReasonId = LookUp_TitleReasons.ReasonId
GROUP BY
  UnitTitleLog.UnitRef,
  UnitTitleLog.UnitRef,
  LookUp_TitleReasons.ReasonDescription
ORDER BY
  UnitTitleLog.UnitRef;
