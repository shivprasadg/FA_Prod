SELECT
  Clients.ClientShNm,
  Units.UnitID,
  Units.UnitUnitNum,
  Schedule.SchFAID
FROM
  (
    (
      SchGrp
      RIGHT JOIN Schedule ON SchGrp.SGrpID = Schedule.SchID
    )
    RIGHT JOIN (
      Clients
      LEFT JOIN MstrLease ON Clients.clientID = MstrLease.Client
    ) ON Schedule.SchMLRef = MstrLease.MLID
  )
  LEFT JOIN (
    Units
    LEFT JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  ) ON SchGrp.SGrpID = Units.GroupREF;
