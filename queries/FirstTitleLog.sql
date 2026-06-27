SELECT
  TL.UnitTitleLogId,
  TL.UnitRef,
  TL.EntryTypeId,
  TL.EntryReasonId,
  TL.TitleNumber,
  TL.TitleSt,
  Nz([EntryReason], '') AS ReasonLogged,
  TL.UnitTitleLogDate,
  TL.ExpectedReturn
FROM
  UnitTitleLog AS TL
  INNER JOIN (
    SELECT
      Min(UnitTitleLog.UnitTitleLogId) AS MinLogId,
      UnitREF
    FROM
      UnitTitleLog
    GROUP BY
      UnitRef
  ) AS T2 ON TL.UnitTitleLogId = T2.MinLogId;
