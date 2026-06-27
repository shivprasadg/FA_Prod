SELECT
  TrailerMilesHours.UnitID,
  TrailerMilesHours.LogActivityDate,
  TrailerMilesHours.MilesLTD,
  TrailerMilesHours.HoursLTD
FROM
  TrailerMilesHours
  INNER JOIN TrailerLastReportDate ON (
    TrailerMilesHours.LogActivityDate = TrailerLastReportDate.MaxOfLogActivityDate
  )
  AND (
    TrailerMilesHours.UnitID = TrailerLastReportDate.UnitID
  )
GROUP BY
  TrailerMilesHours.UnitID,
  TrailerMilesHours.LogActivityDate,
  TrailerMilesHours.MilesLTD,
  TrailerMilesHours.HoursLTD;
