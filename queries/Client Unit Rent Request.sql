SELECT
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.Assignee_Bank,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.ExchangeITGroup,
  SchGrp.ExchangeITMonth,
  SchGrp.LeaseTermNotice,
  Units.unitunitnum,
  Units.unitvin,
  Units.unitsrent,
  Locations.LocationName,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  SchGrp.LeaseExtOption1Term,
  SchGrp.LeaseExtOption1LRF,
  SchGrp.LeaseExtOption1Rent,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Schedule.SchFSL,
  [schfslpmt] AS Expr1
FROM
  (
    (
      Clients
      INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
    )
    INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
  )
  INNER JOIN (
    (
      SchGrp
      INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
    )
    INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
  ) ON Schedule.SchID = SchGrp.SGrpSchID
WHERE
  (
    (
      (Clients.ClientShNm)= [enter short nm]
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  )
ORDER BY
  Schedule.SchID,
  SchGrp.SGrpNO;
