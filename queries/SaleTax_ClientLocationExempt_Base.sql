SELECT
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  Count(vw_SixKeys.AssetID) AS CountOfAssetID,
  IIf(
    [UnitTaxExempt] = False,
    Count([AssetID]),
    0
  ) AS PaysTax,
  IIf(
    [UnitTaxExempt] = True,
    Count([AssetID]),
    0
  ) AS TaxExempt,
  vw_SixKeys.LCity,
  vw_SixKeys.LState,
  [StateSales] + [CountySales] + [CitySales] AS SalesTaxRate,
  Units.[UnitTaxExempt]
FROM
  (
    vw_SixKeys
    LEFT JOIN vw_LocationTaxes ON vw_SixKeys.LocationID = vw_LocationTaxes.LocationID
  )
  INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.LCity,
  vw_SixKeys.LState,
  [StateSales] + [CountySales] + [CitySales],
  Units.[UnitTaxExempt]
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
