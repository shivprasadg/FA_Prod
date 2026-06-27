SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.Status,
  vw_SixKeys.ClientShortName AS SourceCompany,
  Nz([BankShortName], "N/A") AS Assignee,
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.Type,
  vw_SixKeys.MYear,
  vw_SixKeys.Make,
  vw_SixKeys.Model,
  Rmkt.MiParked,
  Rmkt.InspMileage1,
  Locations.LocationName,
  Locations.Locationsn,
  Locations.LCity,
  Locations.LState,
  Locations.LZip
FROM
  (
    (
      vw_SixKeys
      INNER JOIN Rmkt ON vw_SixKeys.AssetID = Rmkt.UnitRef
    )
    INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
  )
  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
WHERE
  (
    (
      (vw_SixKeys.Status)= "TR"
      Or (vw_SixKeys.Status)= "R"
      Or (vw_SixKeys.Status)= "S"
      Or (vw_SixKeys.Status)= "K"
      Or (vw_SixKeys.Status)= "KS"
    )
  );
