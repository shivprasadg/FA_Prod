INSERT INTO TrailerMilesHours (
  Vin, UnitNumber, LogActivityDate,
  miles, hours, ReadDate, UnitID, OBCDataStatus,
  OBCDataDesc, MaintLeaseType, CPM,
  CPH, milesltd, hoursltd, ProviderID
)
SELECT
  vw_SixKeys.Vin,
  vw_SixKeys.Unitnum,
  Costco_DRY.LogActivityDate,
  Costco_DRY.Miles,
  Costco_DRY.Hours,
  DateAdd("d", 33, [LogActivityDate]) AS Expr1,
  vw_SixKeys.UnitID,
  1 AS Expr2,
  "Reporting" AS Expr3,
  "FSL" AS Expr4,
  0 AS Expr6,
  0 AS Expr5,
  Costco_DRY.milesltd,
  Costco_DRY.hoursltd,
  103 AS Expr7
FROM
  vw_SixKeys
  INNER JOIN Costco_DRY ON vw_SixKeys.Unitnum = Costco_DRY.UnitNum;
