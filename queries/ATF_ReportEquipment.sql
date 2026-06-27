SELECT
  ATFUnitSpec.ATFID,
  [UnitModelYear] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubType] & " " & [UnitType] & " " & [Axle] AS UnitDesc,
  ATFUnitSpec.UOM AS [Size],
  ATFUnitSpec.UnitCost,
  ATFUnitSpec.NumUnitsInSpec AS Units,
  ATFUnitSpec.Axle,
  ATFUnitSpec.VendorID,
  Vendors.VendorSNm,
  ATFUnitSpec.EstimatedDeliveryDate,
  ATFUnitSpec.EstimatedBuildDate,
  ATFUnitSpec.SchRent,
  ATFUnitSpec.ATFUnitSpecDesc,
  ATFUnitSpec.ATFGroupSpecID,
  ATFUnitSpec.ATFUnitSpecID,
  Sum(
    CCur(
      Nz([ChildPartCost], 0)
    )
  ) AS TotalPartCost
FROM
  (
    ATFUnitSpec
    LEFT JOIN Vendors ON ATFUnitSpec.VendorID = Vendors.VendorID
  )
  LEFT JOIN ATF_ChildPartCost ON ATFUnitSpec.ATFUnitSpecID = ATF_ChildPartCost.ATFUnitSpecID
GROUP BY
  ATFUnitSpec.ATFID,
  [UnitModelYear] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubType] & " " & [UnitType] & " " & [Axle],
  ATFUnitSpec.UOM,
  ATFUnitSpec.UnitCost,
  ATFUnitSpec.NumUnitsInSpec,
  ATFUnitSpec.Axle,
  ATFUnitSpec.VendorID,
  Vendors.VendorSNm,
  ATFUnitSpec.EstimatedDeliveryDate,
  ATFUnitSpec.EstimatedBuildDate,
  ATFUnitSpec.SchRent,
  ATFUnitSpec.ATFUnitSpecDesc,
  ATFUnitSpec.ATFGroupSpecID,
  ATFUnitSpec.ATFUnitSpecID;
