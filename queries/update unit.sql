UPDATE
  Units
  INNER JOIN UPloadTRACTOR ON Units.UnitID = UPloadTRACTOR.[FA #]
SET
  Units.UnitEngineSerial = [engine #],
  Units.UnitTransSerial = [transmission #],
  Units.UnitFASerial = [front axle #],
  Units.UnitRAFSerial = [rear axle/fwd #],
  Units.UnitRARSerial = [rear axle/rear #],
  Units.UnitBuildDate = [build date],
  Units.UnitInServ = #5/6/2011#,
  Units.UnitInServInd = "P";
