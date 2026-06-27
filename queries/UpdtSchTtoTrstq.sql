INSERT INTO UnitTitleLog (
  UnitRef, EntryTypeId, EntryReasonId,
  UnitTitleLogDate
)
SELECT
  Units.UnitID,
  [forms]![titletotrustq]![out] AS Out,
  [forms]![titletotrustq]![why] AS Why,
  [forms]![titletotrustq]![senddt] AS Sent
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
  );
