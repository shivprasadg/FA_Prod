INSERT INTO TrailerMilesHours (
  Vin, UnitNumber, LogActivityDate,
  miles, hours, ReadDate, UnitID, OBCDataStatus,
  OBCDataDesc, MaintLeaseType, CPM,
  CPH, milesltd, hoursltd, ProviderID
)
SELECT
  vw_SixKeys.Vin,
  vw_SixKeys.Unitnum,
  Costco_Wet.LogActivityDate,
  Costco_Wet.Miles,
  Costco_Wet.Hours,
  DateAdd("d", 33, [LogActivityDate]) AS Expr1,
  vw_SixKeys.UnitID,
  1 AS Expr2,
  "Reporting" AS Expr3,
  "FSL" AS Expr4,
  0 AS Expr6,
  0 AS Expr5,
  Costco_Wet.milesltd,
  Costco_Wet.hoursltd,
  103 AS Expr7
FROM
  vw_SixKeys
  INNER JOIN Costco_Wet ON vw_SixKeys.Unitnum = Costco_Wet.UnitNum
ORDER BY
  vw_SixKeys.Unitnum,
  Costco_Wet.LogActivityDate DESC;
