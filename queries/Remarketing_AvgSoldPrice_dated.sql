SELECT
  LookUp_UnitDesc.DescriptionMed AS Description,
  Avg(Rmkt.SoldPrice) AS AvgSoldPrice,
  Count(vw_SixKeys.AssetID) AS Units
FROM
  (
    vw_SixKeys
    INNER JOIN (
      Units
      INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
    ) ON vw_SixKeys.AssetID = Units.UnitID
  )
  INNER JOIN LookUp_UnitDesc ON vw_SixKeys.AssetID = LookUp_UnitDesc.AssetID
WHERE
  (
    (
      (Rmkt.InvoicePaidDate) Is Null
      Or (Rmkt.InvoicePaidDate)> [forms]![acctmenu]![AsOfDate]
    )
  )
GROUP BY
  LookUp_UnitDesc.DescriptionMed
HAVING
  (
    (
      (
        Avg(Rmkt.SoldPrice)
      )> 0
    )
  );
