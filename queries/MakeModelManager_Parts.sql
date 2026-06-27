SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  [SchNo] & ' |  ' & [SGrpNO] AS SchGrp,
  Units.UnitID,
  Units.UnitUnitNum,
  Units.UnitStatus,
  ChildParts.ChildYear,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  DLookUp(
    "UnitDescShortVer", "vw_SixKeys",
    "AssetID=" & [Units].[UnitID]
  ) AS UnitDesc
FROM
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
    LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
  )
  INNER JOIN MakeModelManagerParts_WorkTable ON Units.UnitID = MakeModelManagerParts_WorkTable.UnitID
ORDER BY
  Clients.ClientShNm,
  [SchNo] & ' |  ' & [SGrpNO],
  Units.UnitID;
