INSERT INTO UnitTitleLog (
  Unitref, UnitTitleLogDate, EntryReasonId,
  EntryTypeId
)
SELECT
  Units.UnitID AS Unitref,
  Rmkt.FATitleBack,
  6 AS reason,
  1 AS EntryTypeId
FROM
  Units
  INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
WHERE
  (
    (
      (Rmkt.FATitleBack) Is Not Null
    )
  )
ORDER BY
  Units.UnitID;
