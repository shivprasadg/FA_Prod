SELECT
  Clients.ClientShNm,
  Recalls.recallref,
  Clients.clientCompanyName,
  Schedule.SchNo,
  Units.unitunitnum,
  Units.unitvin,
  Units.UnitEngineSerial,
  Locations.LocationName,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LCity,
  Locations.LState,
  Locations.LZip,
  MstrLease.MLNo,
  Recalls.RCSent,
  MstrLease.Client,
  Recalls.RCCont,
  RecallCampaign.RecallNo,
  Recalls.RCInspDt,
  Recalls.RCRepairReq,
  Recalls.RCRepDt,
  Recalls.[Repair Done],
  Locations.LocationID
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
  INNER JOIN (
    (
      (
        Units
        INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
      )
      INNER JOIN Recalls ON Units.UnitID = Recalls.RCUnitID
    )
    INNER JOIN RecallCampaign ON Recalls.recallref = RecallCampaign.ID
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Clients.ClientShNm)= forms!recallpform!client
    )
    And (
      (Recalls.recallref)= forms!recallpform!recallno
    )
  )
ORDER BY
  Schedule.SchID,
  Units.unitunitnum,
  Locations.LocationID;
