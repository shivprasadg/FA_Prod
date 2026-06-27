SELECT
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.BDEX,
  vw_SixKeys.Schedule,
  [vw_Units_AutoHaulerPaired].[TractorSubType] & "-AHauler" AS UnitType,
  Count(vw_SixKeys.UnitID) AS Units,
  vw_SixKeys.LeaseTermPrimary AS Term,
  vw_SixKeys.ProjEOL AS ExpiryDate,
  SchGrp.ExchangeITMonth AS ExchangeITTerm,
  Min(
    DateAdd(
      "m", [ExchangeITMonth], [SchGrp].[BLCD]
    )
  ) AS ExchangeITDate,
  Min(
    DateAdd("m",-3, [ProjEOL])
  ) AS NoticeDate,
  vw_Units_AutoHaulerPaired.PairedCost AS OEC,
  CCur(
    [vw_Units_AutoHaulerPaired].[PairedRV]
  ) AS RV
FROM
  (
    (
      (
        (
          (
            vw_SixKeys
            INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
          )
          INNER JOIN UnitTotalOEC ON vw_SixKeys.UnitID = UnitTotalOEC.AssetID
        )
        INNER JOIN Locations ON vw_SixKeys.LocationID = Locations.LocationID
      )
      INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
    )
    INNER JOIN vw_Units_AutoHaulerPaired ON vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired.TractorUnitID
  )
  LEFT JOIN vw_Units_AutoHaulerPaired AS vw_Units_AutoHaulerPaired_1 ON vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired_1.AutoHaulerUnitID
WHERE
  (
    (
      (vw_SixKeys.Status)= "A"
      Or (vw_SixKeys.Status)= "TR"
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (
        vw_Units_AutoHaulerPaired_1.AutoHaulerUnitID
      ) Is Null
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.BDEX,
  vw_SixKeys.Schedule,
  [vw_Units_AutoHaulerPaired].[TractorSubType] & "-AHauler",
  vw_SixKeys.LeaseTermPrimary,
  vw_SixKeys.ProjEOL,
  SchGrp.ExchangeITMonth,
  vw_Units_AutoHaulerPaired.PairedCost,
  CCur(
    [vw_Units_AutoHaulerPaired].[PairedRV]
  ),
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule
ORDER BY
  vw_SixKeys.BDEX,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule;
