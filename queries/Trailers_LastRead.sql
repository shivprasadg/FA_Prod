SELECT
  TrailerMilesHours.UnitID,
  Max(
    TrailerMilesHours.LogActivityDate
  ) AS LastRead,
  TrailerMilesHours.MilesLTD
FROM
  TrailerMilesHours
GROUP BY
  TrailerMilesHours.UnitID,
  TrailerMilesHours.MilesLTD
HAVING
  (
    (
      (
        Max(
          TrailerMilesHours.LogActivityDate
        )
      )< #1/1/2016#
    )
  )
ORDER BY
  Max(
    TrailerMilesHours.LogActivityDate
  ) DESC;
