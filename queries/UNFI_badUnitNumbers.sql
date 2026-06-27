SELECT
  UnfI_Non_Reporting.CompanyName,
  UnfI_Non_Reporting.UnitID AS AssetID,
  FleetStudyImportBase.UnitNumber AS UNFI_UnitNum,
  FleetStudyImportBase.VIN,
  FleetStudyImportBase.UnitYear,
  FleetStudyImportBase.UnitMake,
  FleetStudyImportBase.UnitModel,
  FleetStudyImportBase.UnitType,
  UnfI_Non_Reporting.UnitVIN AS FA_VIN,
  UnfI_Non_Reporting.OBCUnitNumber AS FA_UnitNum
FROM
  UnfI_Non_Reporting
  INNER JOIN FleetStudyImportBase ON UnfI_Non_Reporting.UnitX = FleetStudyImportBase.UnitNumber;
