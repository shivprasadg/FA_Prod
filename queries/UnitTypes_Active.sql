SELECT
  IIf(
    [OffLeaseDate] >= #1/1/2014#
    And [OffLeaseDate] >= #12/31/2014#,
    2014,
    IIf(
      [OffLeaseDate] >= #1/1/2015#
      And [OffLeaseDate] >= #12/31/2015#,
      2015,
      IIf(
        [OffLeaseDate] >= #1/1/2016#
        And [OffLeaseDate] >= #12/31/2016#,
        2016,
        IIf(
          [OffLeaseDate] >= #1/1/2017#
          And [OffLeaseDate] >= #12/31/2017#,
          2017,
          IIf(
            [OffLeaseDate] >= #1/1/2018#
            And [OffLeaseDate] >= #12/31/2018#,
            2018, 1
          )
        )
      )
    )
  ) AS ActiveYear,
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
  Count(vw_sixkeys.AssetID) AS Expr1
FROM
  vw_sixkeys
WHERE
  (
    (
      (vw_sixkeys.MLOrig)= "FA"
    )
  )
GROUP BY
  IIf(
    [OffLeaseDate] >= #1/1/2014#
    And [OffLeaseDate] >= #12/31/2014#,
    2014,
    IIf(
      [OffLeaseDate] >= #1/1/2015#
      And [OffLeaseDate] >= #12/31/2015#,
      2015,
      IIf(
        [OffLeaseDate] >= #1/1/2016#
        And [OffLeaseDate] >= #12/31/2016#,
        2016,
        IIf(
          [OffLeaseDate] >= #1/1/2017#
          And [OffLeaseDate] >= #12/31/2017#,
          2017,
          IIf(
            [OffLeaseDate] >= #1/1/2018#
            And [OffLeaseDate] >= #12/31/2018#,
            2018, 1
          )
        )
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
