SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLNo,
  vw_SixKeys.MLOrig,
  vw_SixKeys.Type,
  Units.UnitStatus,
  Units.UnitGarageLocationRef,
  Locations.LocationName
FROM
  (
    vw_SixKeys
    INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
  )
  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
WHERE
  (
    (
      (vw_SixKeys.clientID)= 4971
    )
    AND (
      (Units.UnitStatus)<> "S"
    )
  );
