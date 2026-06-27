INSERT INTO TrailerMilesHours (
  UnitID, UnitNumber, Miles, ReadDate,
  Hours, LogActivityDate
)
SELECT
  Units.UnitID,
  Units.UnitUnitNum,
  FleetStudyImportBase.Odometer,
  FleetStudyImportBase.OdometerDate,
  FleetStudyImportBase.ReeferHours AS Expr1,
  DatePart('m', [OdometerDate])& "/1/" & DatePart('yyyy', [OdometerDate]) AS LogActivityDate
FROM
  FleetStudyImportBase
  INNER JOIN Units ON FleetStudyImportBase.VIN = Units.UnitVIN;
