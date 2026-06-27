SELECT
  TOP 10 Vendors.VendorSNm AS Vendor,
  Sum(UnitTotalOEC.UnitTotalCost) AS SumOfUnitTotalCost
FROM
  (
    (
      vw_SixKeys
      LEFT JOIN Vendors ON vw_SixKeys.VendorId = Vendors.VendorID
    )
    LEFT JOIN UnitTotalOEC ON vw_SixKeys.AssetID = UnitTotalOEC.AssetID
  )
  INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
WHERE
  (
    (
      (SchGrp.FundDate) Between #3/1/2016#
      And #2/28/2018#
    )
  )
GROUP BY
  Vendors.VendorSNm
ORDER BY
  Sum(UnitTotalOEC.UnitTotalCost) DESC;
