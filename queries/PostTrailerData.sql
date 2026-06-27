INSERT INTO TrailerMilesHours (
  UnitID, UnitNumber, LogActivityDate,
  Miles, MilesLTD, Hours, HoursLTD,
  OBCDataStatus, OBCDataDesc
)
SELECT
  TrailerDataImporter_Maines.UnitID AS Expr1,
  TrailerDataImporter_Maines.UnitUnitNum AS Expr2,
  TrailerDataImporter_Maines.LogActivityDate AS Expr3,
  TrailerDataImporter_Maines.Miles AS Expr4,
  TrailerDataImporter_Maines.MilesLTD AS Expr5,
  TrailerDataImporter_Maines.Hours AS Expr6,
  TrailerDataImporter_Maines.HoursLTD AS Expr7,
  TrailerDataImporter_Maines.OBCDataStatus AS Expr8,
  IIf(
    [ObcDataDesc] = "Reporting,Reporting",
    "Reporting", [ObcDataDesc]
  ) AS DataDesc
FROM
  TrailerDataImporter_Maines;
