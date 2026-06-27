UPDATE
  Units
  INNER JOIN [UpLoad Tractor] ON Units.UnitID = [UpLoad Tractor].[FA #]
SET
  Units.unitunitnum = [lessee unit#],
  Units.unitvin = [vin],
  Units.UnitMSORecd = [mso],
  Units.UnitBuildDate = [build date],
  Units.UnitEstDelDate = [est del date],
  Units.UnitInvoiceNum = [invoice #],
  Units.UnitInvoiceDate = [invoice date],
  Units.UnitEngineSerial = [engine #],
  Units.UnitTransSerial = [transmission #],
  Units.UnitFASerial = [front axle #],
  Units.UnitRAFSerial = [rear axle/fwd #],
  Units.UnitRARSerial = [rear axle/rear #];
