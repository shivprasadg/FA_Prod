SELECT
  Clients.clientCompanyName AS Company,
  RecallCampaign.RecallNo AS RecallNum,
  RecallCampaign.Type,
  RecallCampaign.RecallDesc AS Description,
  RecallCampaign.RecallDt AS NoticeDate,
  Units.UnitUnitNum AS UnitNumber,
  Units.UnitVIN,
  Locations.LocationName,
  Null AS ActionNeeded,
  Recalls.RCRepDt AS ActionCompleted,
  Units.UnitEngineSerial AS EngSerialNo,
  Clients.ClientID,
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
      RecallCampaign
      INNER JOIN (
        (
          Units
          INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
        )
        INNER JOIN Recalls ON Units.UnitID = Recalls.RCUnitID
      ) ON RecallCampaign.ID = Recalls.recallref
    )
    LEFT JOIN Contacts ON Recalls.RCCont = Contacts.contactID
  ) ON SchGrp.SGrpID = Units.GroupREF
ORDER BY
  Locations.LocationName,
  Clients.ClientID,
  Left([unitvin], 6);
