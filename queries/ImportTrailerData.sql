SELECT
  Units.UnitID,
  Units.UnitUnitNum,
  TrailerDataImport.MILRDDT,
  TrailerLastMaxMilesHours.MilesLTD AS LastMilesLTD,
  TrailerDataImport.MILES,
  [MILES] - [LastMilesLTD] AS NewMiles,
  TrailerDataImport.RFRHR,
  TrailerLastMaxMilesHours.HoursLTD AS LastHoursLTD,
  [RFRHR] - [LastHoursLTD] AS NewHours,
  1 AS OBCDataStaus,
  TrailerLastMaxMilesHours.LogActivityDate AS LastLogActivityDate,
  CDate(
    DatePart('m', [MILRDDT])& "/1/" & DatePart('yyyy', [MILRDDT])
  ) AS NewLogActivityDate,
  DateDiff(
    'm', [LastLogActivityDate], [NewLogActivityDate]
  ) AS MonthsBetween,
  CInt([NewMiles] / [MonthsBetween]) AS AvgMiles
FROM
  (
    Units
    INNER JOIN TrailerDataImport ON Units.UnitVIN = TrailerDataImport.VIN
  )
  LEFT JOIN TrailerLastMaxMilesHours ON Units.UnitID = TrailerLastMaxMilesHours.UnitID;
