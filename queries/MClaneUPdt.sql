UPDATE
  (
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
      INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
    )
    INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
  )
  INNER JOIN [Master List] ON Units.UnitUnitNum = [Master List].[IN FA DB?]
SET
  Units.UnitStatus = "TR"
WHERE
  (
    (
      (Clients.ClientShNm)= "mclane"
    )
    AND (
      (MstrLease.MLOrig)<> "fa"
    )
  );
