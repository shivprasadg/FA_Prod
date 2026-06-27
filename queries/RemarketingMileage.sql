SELECT
  TrackedUnitLastMileage.UnitID,
  TrackedUnitLastMileage.MHDate,
  TrackedUnitLastMileage.Odometer,
  Clients.clientobc
FROM
  (
    RemarketingUnitsReport
    INNER JOIN TrackedUnitLastMileage ON RemarketingUnitsReport.AssetID = TrackedUnitLastMileage.UnitID
  )
  LEFT JOIN Clients ON RemarketingUnitsReport.clientID = Clients.clientID;
