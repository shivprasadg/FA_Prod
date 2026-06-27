SELECT
  Non_Reporting_Units.*,
  FleetStudyImportBase.UnitNumber AS UNFI_UnitNum
FROM
  Non_Reporting_Units
  LEFT JOIN FleetStudyImportBase ON Non_Reporting_Units.UnitVIN = FleetStudyImportBase.VIN
WHERE
  (
    (
      (
        FleetStudyImportBase.UnitNumber
      )<> Non_Reporting_Units.UnitUnitNum
    )
    And (
      (FleetStudyImportBase.VIN) Is Not Null
    )
  );
