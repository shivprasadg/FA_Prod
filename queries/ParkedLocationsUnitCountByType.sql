SELECT
  Clients.ClientShNm AS Client,
  Locations.LocationName AS Location,
  Locations.LCity AS City,
  Locations.LState AS State,
  IIf(
    Count([UnitID])> 0,
    Count([UnitID]),
    0
  ) AS Units,
  Sum(
    IIf(
      [Status] = "A"
      Or [Status] = "TR"
      And [MLorig] = "FA",
      1, 0
    )
  ) AS FAUnits,
  Sum(
    IIf([Status] = "B", 1, 0)
  ) AS FANewUnits,
  Sum(
    IIf(
      [SubType] = "Daycab"
      And [Status] = "A"
      And [MLorig] = "FA"
      Or [SubType] = "Daycab"
      And [Status] = "TR"
      And [MLorig] = "FA",
      1, 0
    )
  ) AS DayCabs,
  Sum(
    IIf(
      [SubType] = "Sleeper"
      And [Status] = "A"
      And [MLorig] = "FA"
      Or [SubType] = "Sleeper"
      And [Status] = "TR"
      And [MLorig] = "FA",
      1, 0
    )
  ) AS Sleepers,
  Sum(
    IIf(
      [Type] = "Straight Truck"
      And [Status] = "A"
      And [MLorig] = "FA"
      Or [Type] = "Straight Truck"
      And [Status] = "TR"
      And [MLorig] = "FA",
      1, 0
    )
  ) AS Straight,
  Sum(
    IIf(
      [Type] = "Van"
      And [Status] = "A"
      And [MLorig] = "FA"
      Or [Type] = "Van"
      And [Status] = "TR"
      And [MLorig] = "FA",
      1, 0
    )
  ) AS Van,
  Sum(
    IIf(
      [Type] = "Yard Tractor"
      And [Status] = "A"
      And [MLorig] = "FA"
      Or [Type] = "Yard Trator"
      And [Status] = "TR"
      And [MLorig] = "FA",
      1, 0
    )
  ) AS Yard,
  Sum(
    IIf(
      [Type] = "Trailer"
      And [Status] = "A"
      And [MLorig] = "FA"
      Or [Type] = "Trailer"
      And [Status] = "TR"
      And [MLorig] = "FA",
      1, 0
    )
  ) AS Trailers,
  Min(vw_SixKeys.ActualDeliveryDate) AS FirstDelvery,
  Max(vw_SixKeys.ActualDeliveryDate) AS LastDelvery,
  Max(
    vw_SixKeys.EsitmatedDeliveryDate
  ) AS LastEstDelDate,
  Sum(
    IIf(
      [Status] = "PS"
      Or [Status] = "R", 1, 0
    )
  ) AS RMKTUnits,
  Sum(
    IIf(
      [Status] = "T"
      And [mlorig] <> "FA", 1,
      0
    )
  ) AS TrackingUnits,
  Sum(
    IIf(
      [SubType] = "Daycab"
      And [Status] = "T"
      And [MLorig] <> "FA",
      1, 0
    )
  ) AS T_DayCabs,
  Sum(
    IIf(
      [SubType] = "Sleeper"
      And [Status] = "T"
      And [MLorig] <> "FA",
      1, 0
    )
  ) AS T_Sleepers,
  Sum(
    IIf(
      [Type] = "Straight Truck"
      And [Status] = "T"
      And [MLorig] <> "FA",
      1, 0
    )
  ) AS T_Straight,
  Sum(
    IIf(
      [Type] = "Van"
      And [Status] = "T"
      And [MLorig] <> "FA",
      1, 0
    )
  ) AS T_Van,
  Sum(
    IIf(
      [Type] = "Yard Tractor"
      And [Status] = "T"
      And [MLorig] <> "FA",
      1, 0
    )
  ) AS T_Yard,
  Sum(
    IIf(
      [Type] = "Trailer"
      And [Status] = "T"
      And [MLorig] <> "FA",
      1, 0
    )
  ) AS T_Trailers,
  ClientGroups.ClientGroupName,
  Locations.LocationID,
  Clients.clientID,
  Locations.ClientGroupID
FROM
  (
    Locations
    INNER JOIN (
      Clients
      INNER JOIN vw_SixKeys ON Clients.clientID = vw_SixKeys.clientID
    ) ON Locations.LocationID = vw_SixKeys.LocationID
  )
  INNER JOIN ClientGroups ON Locations.ClientGroupID = ClientGroups.ClientGroupId
GROUP BY
  Clients.ClientShNm,
  Locations.LocationName,
  Locations.LCity,
  Locations.LState,
  ClientGroups.ClientGroupName,
  Locations.LocationID,
  Clients.clientID,
  Locations.ClientGroupID
ORDER BY
  Clients.ClientShNm,
  Locations.LocationName,
  Locations.LCity,
  Locations.LState;
