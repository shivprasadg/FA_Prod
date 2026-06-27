UPDATE
  Units
  INNER JOIN StagingVendorTracking_Lastest ON Units.UnitID = StagingVendorTracking_Lastest.UnitId
SET
  Units.UnitVIN = [StagingVendorTracking_Lastest]![VIN],
  Units.UnitEstDelDate = [StagingVendorTracking_Lastest]![UnitEstDelDate],
  Units.UnitActualDelDate = [StagingVendorTracking_Lastest]![UnitActualDelDate],
  Units.UnitBuildDate = [UnitActualDelDate]![UnitBuildDate],
  Units.UnitPilot = [UnitBuildDate]![IsPilot],
  Units.UnitEngineSerial = [UnitBuildDate]![EngineSer],
  Units.UnitTransSerial = [UnitBuildDate]![TransSer],
  Units.UnitFASerial = [UnitBuildDate]![RAxleFWDSer],
  Units.UnitRAFSerial = [UnitBuildDate]![RAxleREARSer];
