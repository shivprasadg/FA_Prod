SELECT
  Clients.ClientShNm,
  MstrLease.MLOrig,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UNITUNITNUM,
  Units.UNITVIN,
  Units.UnitOriginserv
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
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Clients.ClientShNm)= "UNFI"
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
  )
ORDER BY
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.UNITUNITNUM;
