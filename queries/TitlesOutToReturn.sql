SELECT
  LastTitleLog.UnitTitleLogId,
  LastTitleLog.UnitRef,
  LastTitleLog.EntryTypeId,
  LastTitleLog.EntryReasonId,
  LastTitleLog.TitleNumber,
  LastTitleLog.TitleSt,
  LastTitleLog.ReasonLogged,
  LastTitleLog.UnitTitleLogDate,
  IIf(
    Not IsNull([UnitTitleRecd])
    And [EntryReasonId] = 18,
    False,
    [ExpectedReturn]
  ) AS ToReturn
FROM
  LastTitleLog
  INNER JOIN Units ON LastTitleLog.UnitRef = Units.UnitID
WHERE
  (
    (
      (LastTitleLog.EntryTypeId)= 2
    )
    AND (
      (
        IIf(
          Not IsNull([UnitTitleRecd])
          And [EntryReasonId] = 18,
          False,
          [ExpectedReturn]
        )
      )= True
    )
  )
ORDER BY
  LastTitleLog.UnitTitleLogDate DESC;
