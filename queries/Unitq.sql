UPDATE
  DeliveryUpdate
  INNER JOIN Units ON DeliveryUpdate.FA = Units.UnitID
SET
  Units.unitunitnum = [unit],
  Units.UnitVIN = [vin],
  Units.UnitBuildDate = [build date],
  Units.UnitEstDelDate = [del date],
  Units.UnitPlateNum = [license plate],
  Units.UnitMSORecd = [mso];
