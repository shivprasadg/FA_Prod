SELECT
  Clients.ClientShNm,
  MstrLease.MLOrig,
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  Units.UnitUnitNum,
  Units.UNITVIN,
  Units.UnitID,
  Units.GroupREF,
  Units.UNITYR,
  Units.UNITMAKE,
  ChildParts.UnitREF,
  ChildParts.ChildMake,
  Locations.LocationName,
  Locations.Locationsn,
  Units.UNITFINALCOST,
  ChildParts.childFinalCost,
  Units.UnitEstDelDate
FROM
  (
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
    LEFT JOIN (
      Units
      LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    ) ON SchGrp.SGrpID = Units.GroupREF
  )
  LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
WHERE
  (
    (
      (Clients.ClientShNm) Like "TOY*"
    )
  )
ORDER BY
  Schedule.SchID,
  SchGrp.SGrpNO,
  Units.UnitUnitNum,
  ChildParts.ChildMake;
