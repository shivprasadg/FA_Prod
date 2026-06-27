SELECT
  DatePart(
    "yyyy",
    CDate(
      Nz(
        [OffLeaseDate],
        Nz([ProjEOL], "12/31/2050")
      )
    )
  ) AS EOLYear,
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
  Count(vw_sixkeys.AssetID) AS EOLCount
FROM
  vw_sixkeys
WHERE
  (
    (
      (vw_sixkeys.MLOrig)= "FA"
    )
  )
GROUP BY
  DatePart(
    "yyyy",
    CDate(
      Nz(
        [OffLeaseDate],
        Nz([ProjEOL], "12/31/2050")
      )
    )
  ),
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
  )
HAVING
  (
    (
      (
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
        )
      ) Is Not Null
    )
  )
ORDER BY
  DatePart(
    "yyyy",
    CDate(
      Nz(
        [OffLeaseDate],
        Nz([ProjEOL], "12/31/2050")
      )
    )
  ),
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
  );
