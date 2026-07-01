SELECT
  FS.VIN,
  L.LocationID,
  L.LocationName
FROM
  FleetStudyImport_Work AS FS
  INNER JOIN Locations AS L ON (
    FS.ClientGroupID = L.ClientGroupID
  )
  AND (FS.Location = L.LocationName);
