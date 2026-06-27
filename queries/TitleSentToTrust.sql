INSERT INTO UnitTitleLog (
  UnitRef, UnitTitleLogDate, EntryTypeId,
  EntryReasonId, TitleSt
)
SELECT
  Units.UnitID,
  [forms]![titletotrustq]![SentDate] AS SentDate,
  2 AS Type,
  4 AS Reasonid,
  Units.UnitTitleState
FROM
  (
    Schedule
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Schedule.SchID)= [forms]![titletotrustq]![selsch]
    )
  )
ORDER BY
  Units.UnitID;
