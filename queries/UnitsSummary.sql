SELECT
  Units.UnitEstDelDate,
  Units.UnitUnitNum,
  Units.UnitID,
  Units.GroupREF,
  Units.UnitStatus,
  Units.UnitVIN,
  Units.UnitInvoiceNum,
  Units.UnitFinalCost,
  "Edit" AS EDIT,
  Vendors.VendorName,
  ChildPartSummaryQuery.SumOfchildFinalCost,
  IIf(
    [SumOfchildFinalCost] Is Null, [UnitFinalCost],
    [SumOfchildFinalCost] + [UnitFinalCost]
  ) AS UnitAndChildCost
FROM
  Vendors
  RIGHT JOIN (
    Units
    LEFT JOIN ChildPartSummaryQuery ON Units.UnitID = ChildPartSummaryQuery.UnitREF
  ) ON Vendors.VendorID = Units.VendorREF
ORDER BY
  Units.UnitUnitNum;
