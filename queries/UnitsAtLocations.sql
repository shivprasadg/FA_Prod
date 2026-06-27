SELECT
  vw_ClientFleetList.Client,
  vw_ClientFleetList.DC_Garage,
  Locations.Locationsn,
  Locations.LCity AS City,
  Locations.LState AS State,
  Count(1) AS AllUnitsAt,
  IIf(
    [vw_ClientFleetList].[Status] = "B",
    1, 0
  ) AS NewUnits,
  Min(Units.UnitEstDelDate) AS FirstEstDelvery
FROM
  (
    vw_ClientFleetList
    INNER JOIN Locations ON vw_ClientFleetList.LocationID = Locations.LocationID
  )
  INNER JOIN Units ON vw_ClientFleetList.AssetID = Units.UnitID
WHERE
  (
    (
      (vw_ClientFleetList.Status)<> "O" <> "S" <> "C" <> "KS"
    )
  )
GROUP BY
  vw_ClientFleetList.Client,
  vw_ClientFleetList.DC_Garage,
  Locations.Locationsn,
  Locations.LCity,
  Locations.LState,
  IIf(
    [vw_ClientFleetList].[Status] = "B",
    1, 0
  );
