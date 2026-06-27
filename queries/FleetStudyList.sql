SELECT
  FleetStudyImport.VIN,
  FleetStudyImport.Unitnumber,
  FleetStudyImport.UnitYear,
  FleetStudyImport.UnitMake,
  FleetStudyImport.UnitModel,
  FleetStudyImport.Type,
  FleetStudyImport.Location
FROM
  FleetStudyImport
ORDER BY
  FleetStudyImport.VIN;
