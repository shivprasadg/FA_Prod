SELECT
  Non_Reporting_Units.*,
  Right(
    [UnitUnitNum],
    Len([UnitUnitNum])-1
  ) AS UnitX
FROM
  Non_Reporting_Units
  LEFT JOIN FleetStudyImportBase ON Non_Reporting_Units.UnitVIN = FleetStudyImportBase.VIN
WHERE
  (
    (
      (FleetStudyImportBase.VIN) Is Null
    )
  );
