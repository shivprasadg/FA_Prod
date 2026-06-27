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
  DatePart("q", [ProjEOL])& " Qtr" AS ExpireQtr,
  DatePart("yyyy", [ProjEOL]) AS ExpireYear
FROM
  vw_SixKeys
WHERE
  (
    (
      (vw_SixKeys.UnitDescShortVer) Not Like '%WARRANTY%'
    )
    AND (
      (vw_SixKeys.ProjEOL) Is Not Null
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
  DatePart("q", [ProjEOL])& " Qtr",
  DatePart("yyyy", [ProjEOL]),
  vw_SixKeys.MLOrig
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
