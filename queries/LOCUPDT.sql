SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.UnitVendorPytDate,
  Units.UnitGarageLocationRef,
  Units.UnitTitleLocationRef,
  Units.UnitRegLocationRef
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
      (Schedule.SchNo)= "3"
    )
  );
