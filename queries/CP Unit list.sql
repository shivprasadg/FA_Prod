SELECT
  Clients.ClientShNm,
  MstrLease.MLOrig,
  SchGrp.SGrpNO,
  Units.unitunitnum,
  Locations.LocationID,
  Units.unitvin,
  Locations.LocationName,
  Locations.LCity,
  Locations.LState,
  Units.unitmake,
  Units.UnitAddDt,
  Units.unityr,
  Units.unitinserv,
  Units.UnitOriginser
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
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
WHERE
  (
    (
      (Clients.ClientShNm) Like "Merch*"
    )
    AND (
      (MstrLease.MLOrig)<> "fa"
    )
    AND (
      (Units.unityr)= 1997
      Or (Units.unityr)= 1998
      Or (Units.unityr)= 1999
      Or (Units.unityr)= 2000
      Or (Units.unityr)= 2001
      Or (Units.unityr)= 2002
    )
  )
ORDER BY
  Units.unitunitnum;
