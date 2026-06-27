SELECT
  vw_sixkeys.CompanyName,
  Locations.LocationID,
  Locations.Locationsn,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LZip AS Zip,
  vw_sixkeys.ClientShortName AS Client,
  Count(vw_sixkeys.UnitID) AS Units,
  Locations.Longitude,
  Locations.Latitude,
  Locations.ClientGroupId,
  vw_sixkeys.clientID
FROM
  Locations
  LEFT JOIN vw_sixkeys ON Locations.LocationID = vw_sixkeys.LocationID
GROUP BY
  vw_sixkeys.CompanyName,
  Locations.LocationID,
  Locations.Locationsn,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LZip,
  vw_sixkeys.ClientShortName,
  Locations.Longitude,
  Locations.Latitude,
  Locations.ClientGroupId,
  vw_sixkeys.clientID;
