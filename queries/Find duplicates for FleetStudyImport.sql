SELECT
  FleetStudyImport.[VIN],
  FleetStudyImport.[UnitNumber],
  FleetStudyImport.[UnitYear],
  FleetStudyImport.[UnitMake],
  FleetStudyImport.[UnitModel]
FROM
  FleetStudyImport
WHERE
  (
    (
      (FleetStudyImport.[VIN]) In (
        SELECT
          [VIN]
        FROM
          [FleetStudyImport] As Tmp
        GROUP BY
          [VIN]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  FleetStudyImport.[VIN];
