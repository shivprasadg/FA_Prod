SELECT
  Units.GroupREF,
  Units.UnitID,
  Units.UnitVIN,
  Units.UnitInvoiceNum,
  Units.VendorREF,
  Units.UnitStatus,
  "View/Edit" AS Edit,
  Units.UnitFinalCost,
  Units.UnitBuildDate,
  Units.UnitDeliveryDate AS Expr1
FROM
  Units;
