SELECT
  UnitTitleLog.UnitRef,
  UnitTitleLog.EntryReason,
  UnitTitleLog.EntryReasonId
FROM
  UnitTitleLog
WHERE
  (
    (
      (UnitTitleLog.EntryReasonId)= 18
    )
  )
ORDER BY
  UnitTitleLog.UnitRef;
