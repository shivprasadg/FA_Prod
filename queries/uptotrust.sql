INSERT INTO UnitTitleLog (
  UnitRef, UnitTitleLogDate, EntryTypeId,
  EntryReasonId
)
SELECT
  Units.UnitID,
  [2014].Date,
  [2014].type,
  [2014].reason
FROM
  Units
  INNER JOIN 2014 ON Units.UnitVIN = [2014].VIN;
