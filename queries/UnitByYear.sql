SELECT
  vw_SixKeys.ClientShortName,
  Clients.clientBecameDate AS BecameClient,
  vw_SixKeys.MLOrig,
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
  vw_SixKeys.AssetID,
  vw_SixKeys.Status,
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
  ) AS Accepted,
  vw_SixKeys.BLCD,
  UnitExitOption.EndOrActualDate,
  CDate(
    Nz(
      [UnitExitOption].[EndOrActualDate],
      Nz(
        [OffLeaseDate],
        Nz([ProjEOL], "12/31/2050")
      )
    )
  ) AS OffLease,
  vw_SixKeys.ProjEOL,
  IIf(
    [UnitExitOption].[OptionTypeID] = 1,
    CDate(
      Format(
        [UnitExitOption].[DatePosted], "mm/dd/yyyy"
      )
    )
  ) AS ExchangeIT,
  UnitExitOption.ReplacedByGroupID,
  IIf(
    [UnitExitOption].[OptionTypeID] = 3,
    CDate(
      Format(
        [UnitExitOption].[DatePosted], "mm/dd/yyyy"
      )
    )
  ) AS Extended,
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
  ) AS DYear,
  DatePart("yyyy", [clientBecameDate]) AS BCYear,
  CInt(
    DatePart(
      "yyyy",
      Nz(
        [OffLease],
        CDate("1/1/2050")
      )
    )
  ) AS OYear,
  IIf(
    [BCYear] = 2014
    And [DYear] = [BCYear],
    1, 0
  ) AS New2014,
  IIf(
    [BCYear] = 2015
    And [DYear] = [BCYear],
    1, 0
  ) AS New2015,
  IIf(
    [BCYear] = 2016
    And [DYear] = [BCYear],
    1, 0
  ) AS New2016,
  IIf(
    [BCYear] = 2017
    And [DYear] = [BCYear],
    1, 0
  ) AS New2017,
  IIf(
    [BCYear] = 2018
    And [DYear] = [BCYear],
    1, 0
  ) AS New2018,
  IIf(
    [BCYear] = 2019
    And [DYear] = [BCYear],
    1, 0
  ) AS New2019,
  IIf(
    [DYear] <= 2014
    And [OYear] >= 2014, 1,
    0
  ) AS A2014,
  IIf(
    [DYear] <= 2015
    And [OYear] >= 2015, 1,
    0
  ) AS A2015,
  IIf(
    [DYear] <= 2016
    And [OYear] >= 2016, 1,
    0
  ) AS A2016,
  IIf(
    [DYear] <= 2017
    And [OYear] >= 2017, 1,
    0
  ) AS A2017,
  IIf(
    [DYear] <= 2018
    And [OYear] >= 2018, 1,
    0
  ) AS A2018,
  IIf(
    IsNull(
      [UnitExitOption_1].[ReplacedByGroupID]
    ),
    0,
    1
  ) AS PartOfEXchangeIT,
  IIf(
    [DYear] = 2014
    And [DYear] <> [BCYear],
    1, 0
  ) AS Added2014,
  IIf(
    [DYear] = 2015
    And [DYear] <> [BCYear],
    1, 0
  ) AS Added2015,
  IIf(
    [DYear] = 2016
    And [DYear] <> [BCYear],
    1, 0
  ) AS Added2016,
  IIf(
    [DYear] = 2017
    And [DYear] <> [BCYear],
    1, 0
  ) AS Added2017,
  IIf(
    [DYear] = 2018
    And [DYear] <> [BCYear],
    1, 0
  ) AS Added2018
FROM
  (
    (
      vw_SixKeys
      INNER JOIN Clients ON vw_SixKeys.clientID = Clients.clientID
    )
    LEFT JOIN UnitExitOption ON vw_SixKeys.UnitID = UnitExitOption.UnitID
  )
  LEFT JOIN UnitExitOption AS UnitExitOption_1 ON vw_SixKeys.SchID = UnitExitOption_1.ReplacedByGroupID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
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
  ),
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
  vw_SixKeys.ClientShortName;
