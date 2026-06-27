SELECT
  UnitTitleLog.UnitRef AS AssetId,
  Max(UnitTitleLog.UnitTitleLogDate) AS DateSent
FROM
  UnitTitleLog
GROUP BY
  UnitTitleLog.UnitRef,
  UnitTitleLog.EntryReasonId
HAVING
  (
    (
      (UnitTitleLog.EntryReasonId)= 4
    )
  );
