UPDATE
  Locations
  INNER JOIN LocationsCountyTemp ON Locations.LocationID = LocationsCountyTemp.LocationID
SET
  Locations.LCounty = [LocationsCountyTemp].[LookCounty],
  Locations.LCountyID = [LocationsCountyTemp].[LookCountyID];
