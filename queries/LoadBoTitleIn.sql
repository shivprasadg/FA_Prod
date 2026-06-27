INSERT INTO UnitTitleLog (
  Unitref, UnitTitleLogDate, EntryReasonId,
  EntryTypeId
)
SELECT
  Units.UnitID AS Unitref,
  Rmkt.RmKTBOSRec,
  2 AS reason,
  1 AS EntryTypeId
FROM
  Units
  INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
WHERE
  (
    (
      (Rmkt.RmKTBOSRec) Is Not Null
    )
  )
ORDER BY
  Units.UnitID;
