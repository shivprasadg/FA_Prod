SELECT
  "In" AS [Action],
  UnitTitleLog.UnitRef,
  UnitTitleLog.EntryReason,
  UnitTitleLog.TitleNumber,
  UnitTitleLog.TitleSt,
  UnitTitleLog.UnitTitleLogDate,
  UnitTitleLog.ExpectedReturn,
  UnitTitleLog.UserID,
  UnitTitleLog.LastUpdate,
  UnitTitleLog.UnitTitleLogID,
  UnitTitleLog.EntryTypeId,
  UnitTitleLog.EntryReasonId
FROM
  (
    UnitTitleLog
    LEFT JOIN LookUp_TitleReasons ON UnitTitleLog.EntryReasonId = LookUp_TitleReasons.ReasonId
  )
  INNER JOIN LastTitleLogInOut ON UnitTitleLog.UnitTitleLogID = LastTitleLogInOut.UnitTitleLogID
WHERE
  (
    (
      (LastTitleLogInOut.entrytypeid)= 1
    )
  )
ORDER BY
  UnitTitleLog.UnitRef;
