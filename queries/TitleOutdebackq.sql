SELECT
  UnitTitleLog.UnitRef,
  Last(UnitTitleLog.UnitTitleLogDate) AS LastOfUnitTitleLogDate,
  Last(UnitTitleLog.EntryTypeId) AS LastOfEntryTypeId,
  Last(Reasons.ReasonId) AS LastOfReasonId,
  Reasons.ReasonDescription
FROM
  UnitTitleLog
  INNER JOIN Reasons ON UnitTitleLog.EntryReasonId = Reasons.ReasonId
GROUP BY
  UnitTitleLog.UnitRef,
  Reasons.ReasonDescription
HAVING
  (
    (
      (
        Last(UnitTitleLog.EntryTypeId)
      )= 2
    )
    AND (
      (
        Last(Reasons.ReasonId)
      )= 6
      Or (
        Last(Reasons.ReasonId)
      )= 8
      Or (
        Last(Reasons.ReasonId)
      )= 9
    )
  );
