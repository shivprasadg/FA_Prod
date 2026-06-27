INSERT INTO UnitTitleLog (
  Unitref, UnitTitleLogDate, EntryReasonId,
  TitleNumber, EntryTypeId
)
SELECT
  Units.UnitID AS Unitref,
  Units.UnitTitleRecd AS TitleIn,
  1 AS reason,
  Units.UnitTitleNumber,
  1 AS EntryTypeId
FROM
  Units
WHERE
  (
    (
      (Units.UnitTitleRecd) Is Not Null
    )
  )
ORDER BY
  Units.UnitID;
