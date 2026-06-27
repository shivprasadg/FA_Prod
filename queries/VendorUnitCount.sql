SELECT
  vw_SixKeys.ClientShortName AS Client,
  Vendors.VendorName,
  vw_SixKeys.UnitDescShortVer,
  Count(vw_SixKeys.UnitID) AS UnitCount,
  Sum(vw_SixKeys.Unitcost) AS Unitcost,
  Sum(vw_SixKeys.PartsCost) AS PartsCost,
  Sum(vw_SixKeys.VendorCost) AS VendorCost,
  Vendors.VendorID
FROM
  vw_SixKeys
  INNER JOIN Vendors ON vw_SixKeys.VendorId = Vendors.VendorID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  Vendors.VendorName,
  vw_SixKeys.UnitDescShortVer,
  Vendors.VendorID
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.UnitDescShortVer DESC;
