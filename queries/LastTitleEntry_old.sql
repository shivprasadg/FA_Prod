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
  Last(UnitTitleLog.TitleSt) AS LastOfTitleSt,
  Last(UnitTitleLog.EntryReason) AS LastOfReason,
  UnitTitleLog.ExpectedReturn
FROM
  UnitTitleLog
  LEFT JOIN LookUp_TitleReasons ON UnitTitleLog.EntryReasonId = LookUp_TitleReasons.ReasonId
GROUP BY
  UnitTitleLog.UnitRef,
  UnitTitleLog.ExpectedReturn,
  UnitTitleLog.UnitRef
ORDER BY
  UnitTitleLog.UnitRef;
