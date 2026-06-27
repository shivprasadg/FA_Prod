SELECT
  Clients.ClientShNm,
  MstrLease.Client,
  MstrLease.MLOrig,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UnitUnitNum,
  Units.UnitID,
  Units.UnitStatus,
  Units.UnitAcceptDate,
  Units.UnitInServ,
  Units.UnitInServInd,
  SchGrp.BLCD
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
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF;
