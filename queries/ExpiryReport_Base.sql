SELECT
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.BDEX,
  vw_SixKeys.Schedule,
  Switch(
    [Type] = "Yard Tractor"
    Or [Type] = "Straight Truck"
    Or [Type] = "Van",
    [Type], [SubType] = "Daycab"
    Or [SubType] = "Sleeper"
    Or [SubType] = "Auto Hauler"
    Or [subType] = "Dolly",
    [SubType], [Type] = 'Trailer', [Type],
    [Type] = "Tractor"
    And [SubType] = "TBD",
    "Tractor TBD"
  ) AS UnitType,
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
  Sum(UnitTotalOEC.UnitTotalCost) AS OEC,
  CCur(
    Nz(
      [Units].[UnitSynResid], [Units].[UnitSalesResid]
    )
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
    LEFT JOIN vw_Units_AutoHaulerPaired ON vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired.TractorUnitID
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
        vw_Units_AutoHaulerPaired.TractorUnitID
      ) Is Null
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
  Switch(
    [Type] = "Yard Tractor"
    Or [Type] = "Straight Truck"
    Or [Type] = "Van",
    [Type], [SubType] = "Daycab"
    Or [SubType] = "Sleeper"
    Or [SubType] = "Auto Hauler"
    Or [subType] = "Dolly",
    [SubType], [Type] = 'Trailer', [Type],
    [Type] = "Tractor"
    And [SubType] = "TBD",
    "Tractor TBD"
  ),
  vw_SixKeys.LeaseTermPrimary,
  vw_SixKeys.ProjEOL,
  SchGrp.ExchangeITMonth,
  CCur(
    Nz(
      [Units].[UnitSynResid], [Units].[UnitSalesResid]
    )
  ),
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule
ORDER BY
  vw_SixKeys.BDEX,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule;
