SELECT
  TrailerMilesHours.UnitID,
  Max(
    TrailerMilesHours.LogActivityDate
  ) AS MaxOfLogActivityDate
FROM
  TrailerMilesHours
GROUP BY
  TrailerMilesHours.UnitID;
