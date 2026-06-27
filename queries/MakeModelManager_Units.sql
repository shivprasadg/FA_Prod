SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  [SchNo] & ' |  ' & [SGrpNO] AS SchGrp,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitStatus,
  Units.UnitYr,
  Units.UnitSubType,
  Units.UnitType,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitUOM,
  Units.UnitVin
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
  INNER JOIN MakeModelManager_WorkTable ON Units.UnitID = MakeModelManager_WorkTable.UnitID
ORDER BY
  Clients.ClientShNm,
  [SchNo] & ' |  ' & [SGrpNO],
  Units.UnitID,
  Units.UnitType;
