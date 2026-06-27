SELECT
  TrailerMilesHours.[UnitID],
  TrailerMilesHours.[LogActivityDate],
  TrailerMilesHours.[OBCDataStatus],
  TrailerMilesHours.[Miles],
  TrailerMilesHours.[Hours],
  TrailerMilesHours.ReadDate
FROM
  TrailerMilesHours
WHERE
  (
    (
      (TrailerMilesHours.[UnitID]) In (
        SELECT
          [UnitID]
        FROM
          [TrailerMilesHours] As Tmp
        GROUP BY
          [UnitID],
          [LogActivityDate]
        HAVING
          Count(*)> 1
          And [LogActivityDate] = [TrailerMilesHours].[LogActivityDate]
      )
    )
  )
ORDER BY
  TrailerMilesHours.[UnitID],
  TrailerMilesHours.[LogActivityDate];
