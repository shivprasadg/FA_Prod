SELECT
  DISTINCT Clients.clientshnm
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      (
        (
          Units
          LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
        )
        LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
      )
      LEFT JOIN lookup_GroupMakes ON Units.UnitMake = lookup_GroupMakes.groupMake
    )
    LEFT JOIN Clients AS Clients_1 ON Rmkt.BuyerRef = Clients_1.clientID
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (MstrLease.mlno) Not Like "demo*"
    )
    AND (
      (Units.unitstatus)= "R"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "k"
      Or (Units.unitstatus)= "s"
      Or (Units.unitstatus)= "ks"
    )
    AND (
      (MstrLease.mlorig)= "fa"
      Or (MstrLease.mlorig)= "rmkt"
      Or (MstrLease.mlorig)= "other"
    )
  )
  OR (
    (
      (MstrLease.mlno) Not Like "demo*"
    )
    AND (
      (Units.unitstatus)= "R"
      Or (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "k"
      Or (Units.unitstatus)= "S"
      Or (Units.unitstatus)= "KS"
    )
    AND (
      (MstrLease.mlorig)= "rmkt"
    )
    AND (
      (MstrLease.MLLessor)= "fa llc"
    )
  )
ORDER BY
  Clients.clientshnm;
