SELECT
  Max(UnitTitleLog.UnitTitleLogDate) AS LastLog,
  UnitTitleLog.UnitRef,
  UnitTitleLog.EntryTypeId,
  UnitTitleLog.EntryReasonId,
  UnitTitleLog.TitleNumber,
  UnitTitleLog.UnitTitleZipcodeId,
  UnitTitleLog.TitleSt,
  Nz([Reason], "") AS ReasonOut,
  UnitTitleLog.UnitTitleLogDate,
  UnitTitleLog.ExpectedReturn
FROM
  UnitTitleLog
WHERE
  (
    (
      (UnitTitleLog.EntryTypeId)= 2
    )
  )
GROUP BY
  UnitTitleLog.UnitRef,
  UnitTitleLog.EntryTypeId,
  UnitTitleLog.EntryReasonId,
  UnitTitleLog.TitleNumber,
  UnitTitleLog.UnitTitleZipcodeId,
  UnitTitleLog.TitleSt,
  Nz([Reason], ""),
  UnitTitleLog.UnitTitleLogDate,
  UnitTitleLog.ExpectedReturn
HAVING
  (
    (
      (UnitTitleLog.EntryReasonId)<> 7
    )
  );
