SELECT
  Units.UnitID AS Asset,
  Units.UnitUnitNum,
  Units.VendorREF AS VendorId,
  Units.Unitvin AS VIN,
  Units.UnitBuildDate,
  Units.UnitEstDelDate,
  Units.UnitActualDelDate,
  Units.UnitAcceptDate,
  Units.UnitInServ,
  Units.UnitDeliveryLocationRef,
  Units.UnitStatus,
  Units.UnitType,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel
FROM
  Units
  INNER JOIN Bulk_Update_WorkTable ON Units.UnitID = Bulk_Update_WorkTable.UnitID
ORDER BY
  Units.UnitID;
