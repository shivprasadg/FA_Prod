SELECT
  Mileage.UnitID,
  Max(Mileage.LogActivityDate) AS LastOfMHDate,
  Max(Mileage.OdometerEnd) AS Odometer,
  lookup_groupTypes.grpClass,
  Clients.clientobc,
  Units.UnitExckPortRpt,
  Units.UnitID AS UnitRef
FROM
  (
    Clients
    INNER JOIN (
      (
        MstrLease
        INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
      )
      INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
    ) ON Clients.clientID = MstrLease.Client
  )
  INNER JOIN (
    (
      Units
      LEFT JOIN Mileage ON Units.UnitID = Mileage.UnitID
    )
    INNER JOIN lookup_groupTypes ON Units.UnitType = lookup_groupTypes.groupType
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Mileage.UnitID) Is Not Null
    )
  )
GROUP BY
  Mileage.UnitID,
  lookup_groupTypes.grpClass,
  Clients.clientobc,
  Units.UnitExckPortRpt,
  Units.UnitID
HAVING
  (
    (
      (lookup_groupTypes.grpClass)<> "trailer"
    )
    AND (
      (Clients.clientobc)<> "none"
    )
    AND (
      (Units.UnitExckPortRpt)= False
    )
  )
ORDER BY
  Mileage.UnitID;
