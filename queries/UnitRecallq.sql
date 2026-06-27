SELECT
  Clients.clientCompanyName AS Company,
  RecallCampaign.RecallNo,
  RecallCampaign.Type,
  RecallCampaign.RecallDesc AS Descsription,
  RecallCampaign.RecallDt AS [Recall Dt],
  RecallCampaign.RecallRec AS [RA Received],
  Units.UnitUnitNum AS [Unit No],
  Units.unitvin AS VIIN,
  Units.UnitEngineSerial AS [Eng Serioa NO],
  Locations.LocationName,
  Locations.LocationStreet AS Add1,
  Locations.LAddrLine1 AS Add2,
  Locations.LCity,
  Locations.LState AS St,
  Locations.LZip AS ZiP,
  Schedule.SchNo,
  Null AS [Inspected (date)],
  Null AS [Repair Needed (Y/N)],
  Null AS [Repair Made(dt)],
  Contacts.contactFullName,
  Contacts.contactPrimaryPhone,
  Contacts.contactEmail,
  Units.UnitID,
  Recalls.RCSent
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
        (
          Units
          LEFT JOIN Recalls ON Units.UnitID = Recalls.RCUnitID
        )
        INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
      )
      LEFT JOIN RecallCampaign ON Recalls.recallref = RecallCampaign.ID
    )
    LEFT JOIN Contacts ON Recalls.RCCont = Contacts.contactID
  ) ON SchGrp.SGrpID = Units.GroupREF
ORDER BY
  RecallCampaign.RecallDt DESC,
  Locations.LocationName,
  MstrLease.Client,
  Left([unitvin], 6);
