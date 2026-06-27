SELECT
  Last(UnitTitleLog.UnitTitleLogId) AS TitleLogId,
  UnitTitleLog.UnitRef AS AssetId,
  Max(UnitTitleLog.UnitTitleLogDate) AS LogDate,
  Last(
    IIf([entrytypeid] = 1, "In", "Out")
  ) AS [Action],
  Last(Reasons.ReasonDescription) AS Description,
  Last(UnitTitleLog.TitleNumber) AS TitleNumber,
  Last(UnitTitleLog.ExpectedReturn) AS ExpectedReturn
FROM
  UnitTitleLog
  LEFT JOIN Reasons ON UnitTitleLog.EntryReasonId = Reasons.ReasonId
GROUP BY
  UnitTitleLog.UnitRef
ORDER BY
  UnitTitleLog.UnitRef DESC;
