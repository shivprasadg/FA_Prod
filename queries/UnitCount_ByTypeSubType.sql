SELECT
  IIf([MLOrig] = "FA", "FA", "Tracking") AS Leased,
  Nz(
    [DealYear],
    IIf(
      [mlorig] = "FA",
      DatePart(
        "yyyy",
        Nz(
          [BuildDate],
          Nz(
            [ReportStartDate], [EsitmatedDeliveryDate]
          )
        )
      ),
      [MYear] - 1
    )
  ) AS FAYear,
  vw_SixKeys.Status,
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
  Count(vw_SixKeys.UnitID) AS Units
FROM
  vw_SixKeys
  LEFT JOIN Awarded_Deals_Client ON (
    vw_SixKeys.clientID = Awarded_Deals_Client.clientID
  )
  AND (
    vw_SixKeys.FAID = Awarded_Deals_Client.FAID
  )
GROUP BY
  IIf([MLOrig] = "FA", "FA", "Tracking"),
  Nz(
    [DealYear],
    IIf(
      [mlorig] = "FA",
      DatePart(
        "yyyy",
        Nz(
          [BuildDate],
          Nz(
            [ReportStartDate], [EsitmatedDeliveryDate]
          )
        )
      ),
      [MYear] - 1
    )
  ),
  vw_SixKeys.Status,
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
      )<> ""
    )
  )
ORDER BY
  IIf([MLOrig] = "FA", "FA", "Tracking"),
  Nz(
    [DealYear],
    IIf(
      [mlorig] = "FA",
      DatePart(
        "yyyy",
        Nz(
          [BuildDate],
          Nz(
            [ReportStartDate], [EsitmatedDeliveryDate]
          )
        )
      ),
      [MYear] - 1
    )
  ) DESC,
  vw_SixKeys.Status,
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
