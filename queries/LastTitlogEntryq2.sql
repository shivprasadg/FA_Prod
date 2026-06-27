SELECT
  LastTitlogEntryq1.UnitRef,
  UnitTitleLog.EntryTypeId,
  UnitTitleLog.EntryReasonId,
  LastTitlogEntryq1.MaxOfUnitTitleLogId,
  Units.UNITVIN
FROM
  Units
  INNER JOIN (
    LastTitlogEntryq1
    INNER JOIN UnitTitleLog ON (
      LastTitlogEntryq1.MaxOfUnitTitleLogId = UnitTitleLog.UnitTitleLogId
    )
    AND (
      LastTitlogEntryq1.MaxOfUnitTitleLogDate = UnitTitleLog.UnitTitleLogDate
    )
  ) ON Units.UnitID = LastTitlogEntryq1.UnitRef
WHERE
  (
    (
      (UnitTitleLog.EntryTypeId)= 2
    )
    AND (
      (UnitTitleLog.EntryReasonId)<> 4
      And (UnitTitleLog.EntryReasonId)<> 7
    )
  );
