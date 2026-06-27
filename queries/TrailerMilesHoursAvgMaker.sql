SELECT
  Units.UnitID,
  Units.UnitUnitNum,
  FleetStudyImportBase.VIN,
  Units.UnitModel,
  Units.UnitMake,
  Units.UnitType,
  Units.UnitInServ,
  FleetStudyImportBase.OdometerDate AS ReadDate,
  [FleetStudyImportBase].[ReeferHours] - [HoursLTD] AS ReeferHours,
  [FleetStudyImportBase].[Odometer] - [MilesLTD] AS TrailerMiles,
  DateDiff(
    "m", [UnitInServ], [OdometerDate]
  ) AS Months,
  Round([Odometer] / [Months], 1) AS AvgMilesPM,
  Round([ReeferHours] / [months], 1) AS AvgHoursPM,
  DatePart('m', [OdometerDate])& "/1/" & DatePart('yyyy', [OdometerDate]) AS LogActivityDate
FROM
  (
    Units
    INNER JOIN FleetStudyImportBase ON Units.UnitVIN = FleetStudyImportBase.VIN
  )
  LEFT JOIN TrailerLastMaxMilesHours ON Units.UnitID = TrailerLastMaxMilesHours.UnitID;
