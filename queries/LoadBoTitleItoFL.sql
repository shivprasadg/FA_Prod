INSERT INTO UnitTitleLog (
  Unitref, UnitTitleLogDate, EntryReasonId,
  EntryTypeId
)
SELECT
  Units.UnitID AS Unitref,
  Rmkt.TitleOut,
  5 AS reason,
  2 AS EntryTypeId
FROM
  Units
  INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
WHERE
  (
    (
      (Rmkt.TitleOut) Is Not Null
    )
  )
ORDER BY
  Units.UnitID;
