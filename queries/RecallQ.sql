SELECT
  Recalls.RCID,
  [rcdt] & [rcno] & [rcdesc] AS Grp,
  Recalls.RCDt,
  Recalls.RCNo,
  Recalls.RCDesc,
  Clients.ClientShNm,
  Locations.LocationName,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.Assignee_Bank,
  Units.UnitUnitNum,
  Units.unitvin,
  Recalls.RCMake,
  Recalls.RCCompMake,
  Recalls.RCRec,
  Recalls.RCSent,
  Recalls.RCConf,
  IIf(
    IsNull([rcsent]),
    1,
    IIf(
      IsNull([rcconf]),
      2,
      IIf(
        IsNull([rcinspdt]),
        3,
        IIf(
          IsNull([rcrepairreq]),
          4,
          IIf(
            IsNull([rcrepdt]),
            5,
            "ERR"
          )
        )
      )
    )
  ) AS ROrder,
  Recalls.RCInspDt,
  Recalls.rcrepairreq,
  Recalls.rcrepdt,
  Locations.Locationsn
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
    INNER JOIN Recalls ON Units.UnitID = Recalls.RCUnitID
  )
  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
WHERE
  (
    (
      (Recalls.RCSent) Is Null
    )
  )
  OR (
    (
      (Recalls.RCConf) Is Null
    )
  )
  OR (
    (
      (Recalls.RCInspDt)= "isnull"
    )
    AND (
      (Recalls.rcrepairreq) Is Null
    )
  )
  OR (
    (
      (Recalls.rcrepairreq)= Yes
    )
    AND (
      (Recalls.rcrepdt) Is Null
    )
  );
