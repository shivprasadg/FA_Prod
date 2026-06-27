SELECT
  DatePart(
    "yyyy",
    CDate(
      Nz(
        Nz(
          [AcceptedDate],
          Nz(
            [ActualDeliveryDate], [EsitmatedDeliveryDate]
          )
        ),
        Date()
      )
    )
  ) AS Accepted,
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
  Count(vw_sixkeys.AssetID) AS UnitAccepted
FROM
  vw_sixkeys
  INNER JOIN Clients ON vw_sixkeys.clientID = Clients.clientID
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
        Nz(
          [AcceptedDate],
          Nz(
            [ActualDeliveryDate], [EsitmatedDeliveryDate]
          )
        ),
        Date()
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
        Nz(
          [AcceptedDate],
          Nz(
            [ActualDeliveryDate], [EsitmatedDeliveryDate]
          )
        ),
        Date()
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
