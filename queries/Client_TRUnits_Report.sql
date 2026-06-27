SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  Units.UnitID AS AssetId,
  Units.UnitUnitNum AS UnitNo,
  Units.Unitvin AS VIN,
  Units.UnitStatus,
  Units.UnitYr AS ModelYear,
  Units.UnitMake AS Make,
  Units.unittype AS Type,
  Units.PakrDt AS Parked,
  Locations.LocationName AS Garage,
  CDbl([MaxOfMileage]) AS ReportedMileage,
  CDate([MaxOfMHDate]) AS LastReportedDate
FROM
  (
    (
      vw_SixKeys
      INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
    )
    INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
  )
  INNER JOIN MaxMileageq ON Units.UnitID = MaxMileageq.UnitRef
WHERE
  (
    (
      (vw_SixKeys.clientID)= 4971
    )
    AND (
      (Units.UnitStatus)= "TR"
    )
  )
ORDER BY
  Units.UnitUnitNum;
