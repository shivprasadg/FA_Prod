SELECT
  vw_SixKeys.ClientShortName,
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
  Count(vw_SixKeys.AssetID) AS Units,
  DatePart("q", [InvoicePaidDate])& " Qtr" AS SoldQtr,
  DatePart("yyyy", [InvoicePaidDate]) AS SoldYear
FROM
  (
    vw_SixKeys
    LEFT JOIN vw_Units_AutoHaulerPaired ON vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired.TractorUnitID
  )
  INNER JOIN Rmkt ON vw_SixKeys.UnitID = Rmkt.UnitRef
WHERE
  (
    (
      (vw_SixKeys.UnitDescShortVer) Not Like '%WARRANTY%'
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
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
  DatePart("q", [InvoicePaidDate])& " Qtr",
  DatePart("yyyy", [InvoicePaidDate])
HAVING
  (
    (
      (
        DatePart("yyyy", [InvoicePaidDate])
      ) Is Not Null
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
