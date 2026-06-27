UPDATE
  Units
  INNER JOIN [UPLOAD TRAILER] ON Units.UnitID = [UPLOAD TRAILER].[FA Unit]
SET
  Units.unitunitnum = [lessee unit #],
  Units.unitvin = [vin],
  Units.UnitPlateNum = [license plate #],
  Units.UnitMSORecd = [mso],
  Units.UnitFinalCost = [cost],
  Units.UnitBuildDate = [build date],
  Units.UnitEstDelDate = [est del date],
  Units.UnitInvoiceNum = [invoice #],
  Units.UnitInvoiceDate = [invoice date];
