SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  MstrLease.MLOrig,
  vw_SixKeys.SGrpID,
  Clients.clientID,
  Schedule.SchID,
  SchGrp.SGrpID,
  Units.UnitID,
  Clients.clientID
FROM
  Units
  RIGHT JOIN (
    (
      (
        (
          SchGrp
          LEFT JOIN vw_SixKeys ON SchGrp.SGrpID = vw_SixKeys.SGrpID
        )
        LEFT JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
      )
      LEFT JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID
    )
    LEFT JOIN Clients ON MstrLease.Client = Clients.ClientGroupId
  ) ON Units.GroupREF = SchGrp.SGrpID
WHERE
  (
    (
      (vw_SixKeys.SGrpID) Is Null
    )
  );
