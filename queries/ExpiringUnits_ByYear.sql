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
  DatePart(
    "yyyy",
    CDate(
      Nz(
        [OffLeaseDate],
        Nz([ProjEOL], "12/31/2050")
      )
    )
  ) AS OYear,
  IIf([OYear] <= 2013, 1, 0) AS O2013,
  IIf([OYear] <= 2014, 1, 0) AS O2014,
  IIf([OYear] = 2015, 1, 0) AS O2015,
  IIf([OYear] = 2016, 1, 0) AS O2016,
  IIf([OYear] = 2017, 1, 0) AS O2017,
  IIf([OYear] = 2018, 1, 0) AS O2018,
  IIf([OYear] = 2019, 1, 0) AS O2019
FROM
  vw_SixKeys
  INNER JOIN Clients ON vw_SixKeys.clientID = Clients.clientID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
