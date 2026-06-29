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
  INNER JOIN [SELECT Max(UnitTitleLog].[UnitTitleLogDate) AS MaxLogDate, UnitREF FROM UnitTitleLog GROUP BY UnitRef] AS T2 ON (
    TL.UnitTitleLogDate = T2.MaxLogDate
  )
  AND (TL.UnitRef = T2.UnitREF)
ORDER BY
  TL.UnitTitleLogDate DESC;
