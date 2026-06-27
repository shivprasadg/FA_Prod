SELECT
  SchGrp.MTMRate,
  SchGrp.MTMRent,
  SchGrp.SGrpNO,
  Schedule.SchNo,
  MstrLease.MLOrig,
  Clients.ClientShNm,
  Units.unitfinalcost,
  Units.MTMRent,
  Units.FirstExtRent
FROM
  Clients
  INNER JOIN (
    MstrLease
    INNER JOIN (
      Schedule
      INNER JOIN (
        SchGrp
        INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
      ) ON Schedule.SchID = SchGrp.SGrpSchID
    ) ON MstrLease.MLID = Schedule.SchMLRef
  ) ON Clients.clientID = MstrLease.Client
WHERE
  (
    (
      (SchGrp.MTMRate)> 0
    )
  )
  OR (
    (
      (SchGrp.MTMRent)> 0
    )
  )
ORDER BY
  SchGrp.SGrpNO,
  Schedule.SchNo,
  Units.unitfinalcost;
