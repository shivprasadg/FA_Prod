INSERT INTO UnitTitleLog (
  Unitref, UnitTitleLogDate, EntryReasonId,
  EntryTypeId
)
SELECT
  Units.UnitID AS Unitref,
  Rmkt.RmktBOSSent,
  7 AS reason,
  2 AS EntryTypeId
FROM
  Units
  INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
WHERE
  (
    (
      (Rmkt.RmktBOSSent) Is Not Null
    )
  )
ORDER BY
  Units.UnitID;
