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
  IIf(
    [OptionTypeID] = 1,
    CInt(
      DatePart(
        "yyyy",
        CDate(
          Nz(
            Nz(
              [schgrp].[BLCD],
              Nz(
                [AcceptedDate], [ActualDeliveryDate]
              )
            ),
            Date()
          )
        )
      )
    )
  ) AS XITYear,
  Nz(
    [Extended_ByYear].[EXTendYear],
    0
  ) AS EXTendYear,
  IIf(
    Abs(
      DateDiff(
        "m", [clientBecameDate], [schgrp].[BLCD]
      )
    )<= 12,
    [Dyear]
  ) AS BCYear,
  IIf(
    [status] = "B",
    CDate("1/31/2020"),
    CInt(
      DatePart(
        "yyyy",
        CDate(
          Nz(
            Nz(
              [funddate],
              Nz(
                [schgrp].[BLCD],
                Nz(
                  [AcceptedDate], [ActualDeliveryDate]
                )
              )
            ),
            "1/31/2020"
          )
        )
      )
    )
  ) AS DYear,
  CInt(
    DatePart(
      "yyyy",
      CDate(
        Nz(
          [OffLeaseDate],
          Nz([ProjEOL], "12/31/2050")
        )
      )
    )
  ) AS OYear,
  IIf(
    [Dyear] <= 2012
    And [OYear] > 2012, 1,
    0
  ) AS A2012,
  IIf(
    [Dyear] <= 2013
    And [OYear] > 2013, 1,
    0
  ) AS A2013,
  IIf(
    [Dyear] <= 2014
    And [OYear] > 2014, 1,
    0
  ) AS A2014,
  IIf(
    [Dyear] <= 2015
    And [OYear] > 2015, 1,
    0
  ) AS A2015,
  IIf(
    [Dyear] <= 2016
    And [OYear] > 2016, 1,
    0
  ) AS A2016,
  IIf(
    [Dyear] <= 2017
    And [OYear] > 2017, 1,
    0
  ) AS A2017,
  IIf(
    [Dyear] <= 2018
    And [OYear] > 2018, 1,
    0
  ) AS A2018,
  IIf([DYear] = 2013, 1, 0) AS D2013,
  IIf([DYear] = 2014, 1, 0) AS D2014,
  IIf([DYear] = 2015, 1, 0) AS D2015,
  IIf([DYear] = 2016, 1, 0) AS D2016,
  IIf([DYear] = 2017, 1, 0) AS D2017,
  IIf([DYear] = 2018, 1, 0) AS D2018,
  IIf([OYear] = 2013, 1, 0) AS O2013,
  IIf([OYear] = 2014, 1, 0) AS O2014,
  IIf([OYear] = 2015, 1, 0) AS O2015,
  IIf([OYear] = 2016, 1, 0) AS O2016,
  IIf([OYear] = 2017, 1, 0) AS O2017,
  IIf([OYear] = 2018, 1, 0) AS O2018,
  IIf([OYear] = 2019, 1, 0) AS O2019,
  IIf([BCYear] = 2013, 1, 0) AS N2013,
  IIf([BCYear] = 2014, 1, 0) AS N2014,
  IIf([BCYear] = 2015, 1, 0) AS N2015,
  IIf([BCYear] = 2016, 1, 0) AS N2016,
  IIf([BCYear] = 2017, 1, 0) AS N2017,
  IIf([BCYear] = 2018, 1, 0) AS N2018,
  0 AS XIT2013,
  IIf([XITYear] = 2014, 1, 0) AS XIT2014,
  IIf([XITYear] = 2015, 1, 0) AS XIT2015,
  IIf([XITYear] = 2016, 1, 0) AS XIT2016,
  IIf([XITYear] = 2017, 1, 0) AS XIT2017,
  IIf([XITYear] = 2018, 1, 0) AS XIT2018,
  0 AS EYear2013,
  0 AS EYear2014,
  0 AS EYear2015,
  0 AS EYear2016,
  IIf([EXTendYear] = 2017, 1, 0) AS EYear2017,
  IIf([EXTendYear] = 2018, 1, 0) AS EYear2018,
  IIf([Status] = "B", 1, 0) AS B2018,
  IIf([Status] = "TR", 1, 0) AS TR2018
FROM
  (
    (
      (
        (
          vw_SixKeys
          INNER JOIN Clients ON vw_SixKeys.clientID = Clients.clientID
        )
        LEFT JOIN EXchangeIT_ByYear ON vw_SixKeys.SGrpID = EXchangeIT_ByYear.ReplacedByGroupID
      )
      LEFT JOIN Extended_ByYear ON vw_SixKeys.AssetID = Extended_ByYear.UnitID
    )
    LEFT JOIN vw_Units_AutoHaulerPaired ON vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired.TractorUnitID
  )
  INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (vw_SixKeys.Status)= "B"
    )
    AND (
      (
        vw_Units_AutoHaulerPaired.TractorUnitID
      ) Is Null
    )
  )
ORDER BY
  IIf(
    [status] = "B",
    CDate("1/31/2020"),
    CInt(
      DatePart(
        "yyyy",
        CDate(
          Nz(
            Nz(
              [funddate],
              Nz(
                [schgrp].[BLCD],
                Nz(
                  [AcceptedDate], [ActualDeliveryDate]
                )
              )
            ),
            "1/31/2020"
          )
        )
      )
    )
  ),
  vw_SixKeys.ClientShortName;
